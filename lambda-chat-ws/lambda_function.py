import json
import boto3
import os
import time
import datetime
import PyPDF2
import csv
import re
import traceback
import base64
import operator
import requests

from botocore.config import Config
from io import BytesIO
from urllib import parse
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain.docstore.document import Document
from langchain_core.prompts import MessagesPlaceholder, ChatPromptTemplate
from langchain.memory import ConversationBufferWindowMemory
from langchain_aws import ChatBedrock
from multiprocessing import Process, Pipe
from PIL import Image
from langchain_core.messages import BaseMessage, HumanMessage, AIMessage, SystemMessage, ToolMessage
from langgraph.graph import START, END, StateGraph
from pydantic.v1 import BaseModel, Field
from typing import Annotated, List, Tuple, TypedDict, Literal, Sequence, Union
from tavily import TavilyClient  
from langgraph.constants import Send
from opensearchpy import OpenSearch, RequestsHttpConnection, AWSV4SignerAuth
from langchain_community.vectorstores.opensearch_vector_search import OpenSearchVectorSearch
from langchain_aws import BedrockEmbeddings
from langchain_community.vectorstores.faiss import FAISS

from langgraph.graph.message import add_messages
from langgraph.prebuilt import ToolNode
from langchain.agents import tool
from bs4 import BeautifulSoup
from pytz import timezone

s3 = boto3.client('s3')
s3_bucket = os.environ.get('s3_bucket') # bucket name
s3_prefix = os.environ.get('s3_prefix')
callLogTableName = os.environ.get('callLogTableName')
path = os.environ.get('path')
doc_prefix = s3_prefix+'/'
debugMessageMode = os.environ.get('debugMessageMode', 'false')
opensearch_account = os.environ.get('opensearch_account')
opensearch_passwd = os.environ.get('opensearch_passwd')
opensearch_url = os.environ.get('opensearch_url')
enableParentDocumentRetrival = os.environ.get('enableParentDocumentRetrival')
vectorIndexName = os.environ.get('vectorIndexName')
index_name = vectorIndexName

projectName = os.environ.get('projectName')
LLM_for_chat = json.loads(os.environ.get('LLM_for_chat'))
LLM_for_multimodal = json.loads(os.environ.get('LLM_for_multimodal'))
LLM_embedding = json.loads(os.environ.get('LLM_embedding'))
selected_chat = 0
selected_multimodal = 0
selected_embedding = 0
useEnhancedSearch = False
enableHybridSearch = os.environ.get('enableHybridSearch')
minDocSimilarity = 250
grade_state = "LLM" # LLM, PRIORITY_SEARCH, OTHERS
numberOfDocs = 2
    
multi_region_models = [   # claude sonnet 3.0
    {   
        "bedrock_region": "us-west-2", # Oregon
        "model_type": "claude3",
        "max_tokens": 4096,
        "model_id": "anthropic.claude-3-sonnet-20240229-v1:0"
    },
    {
        "bedrock_region": "us-east-1", # N.Virginia
        "model_type": "claude3",
        "max_tokens": 4096,
        "model_id": "anthropic.claude-3-sonnet-20240229-v1:0"
    },
    {
        "bedrock_region": "ca-central-1", # Canada
        "model_type": "claude3",
        "max_tokens": 4096,
        "model_id": "anthropic.claude-3-sonnet-20240229-v1:0"
    },
    {
        "bedrock_region": "eu-west-2", # London
        "model_type": "claude3",
        "max_tokens": 4096,
        "model_id": "anthropic.claude-3-sonnet-20240229-v1:0"
    },
    {
        "bedrock_region": "sa-east-1", # Sao Paulo
        "model_type": "claude3",
        "max_tokens": 4096,
        "model_id": "anthropic.claude-3-sonnet-20240229-v1:0"
    }
]
multi_region = 'disable'

titan_embedding_v1 = [  # dimension = 1536
  {
    "bedrock_region": "us-west-2", # Oregon
    "model_type": "titan",
    "model_id": "amazon.titan-embed-text-v1"
  },
  {
    "bedrock_region": "us-east-1", # N.Virginia
    "model_type": "titan",
    "model_id": "amazon.titan-embed-text-v1"
  }
]
priority_search_embedding = titan_embedding_v1
selected_ps_embedding = 0

reference_docs = []

tavily_api_key = ""
weather_api_key = ""
def load_secrets():
    global tavily_api_key, weather_api_key
    secretsmanager = boto3.client('secretsmanager')
    
    # api key to use LangSmith
    langsmith_api_key = ""
    try:
        get_langsmith_api_secret = secretsmanager.get_secret_value(
            SecretId=f"langsmithapikey-{projectName}"
        )
        # print('get_langsmith_api_secret: ', get_langsmith_api_secret)
        
        secret = json.loads(get_langsmith_api_secret['SecretString'])
        #print('secret: ', secret)
        langsmith_api_key = secret['langsmith_api_key']
        langchain_project = secret['langchain_project']
    except Exception as e:
        raise e

    if langsmith_api_key:
        os.environ["LANGCHAIN_API_KEY"] = langsmith_api_key
        os.environ["LANGCHAIN_TRACING_V2"] = "true"
        os.environ["LANGCHAIN_PROJECT"] = langchain_project
        
    # api key to use Tavily Search    
    try:
        get_tavily_api_secret = secretsmanager.get_secret_value(
            SecretId=f"tavilyapikey-{projectName}"
        )
        #print('get_tavily_api_secret: ', get_tavily_api_secret)
        
        secret = json.loads(get_tavily_api_secret['SecretString'])
        # print('secret: ', secret)
        if secret['tavily_api_key']:
            tavily_api_key = json.loads(secret['tavily_api_key'])
            # print('tavily_api_key: ', tavily_api_key)
    except Exception as e: 
        raise e
    
    try:
        get_weather_api_secret = secretsmanager.get_secret_value(
            SecretId=f"openweathermap-{projectName}"
        )
        #print('get_weather_api_secret: ', get_weather_api_secret)
        
        secret = json.loads(get_weather_api_secret['SecretString'])
        #print('secret: ', secret)
        weather_api_key = secret['weather_api_key']

    except Exception as e:
        raise e
load_secrets()

def check_tavily_secret(tavily_api_key):
    query = 'what is LangGraph'
    valid_keys = []
    for key in tavily_api_key:
        try:
            tavily_client = TavilyClient(api_key=key)
            response = tavily_client.search(query, max_results=1)
            # print('tavily response: ', response)
            
            if 'results' in response and len(response['results']):
                valid_keys.append(key)
        except Exception as e:
            print('Exception: ', e)
    # print('valid_keys: ', valid_keys)
    
    return valid_keys

#tavily_api_key = check_tavily_secret(tavily_api_key)
#print('The number of valid tavily api keys: ', len(tavily_api_key))

selected_tavily = -1
if len(tavily_api_key):
    os.environ["TAVILY_API_KEY"] = tavily_api_key[0]
    selected_tavily = 0
   
def tavily_search(query, k):
    global selected_tavily
    docs = []
        
    if selected_tavily != -1:
        selected_tavily = selected_tavily + 1
        if selected_tavily == len(tavily_api_key):
            selected_tavily = 0

        try:
            tavily_client = TavilyClient(api_key=tavily_api_key[selected_tavily])
            response = tavily_client.search(query, max_results=k)
            # print('tavily response: ', response)
            
            if "url" in r:
                url = r.get("url")
                
            for r in response["results"]:
                name = r.get("title")
                if name is None:
                    name = 'WWW'
            
                docs.append(
                    Document(
                        page_content=r.get("content"),
                        metadata={
                            'name': name,
                            'url': url,
                            'from': 'tavily'
                        },
                    )
                )   
        except Exception as e:
            print('Exception: ', e)
    return docs

# result = tavily_search('what is LangChain', 2)
# print('search result: ', result)

def refresh_opensearch_index():
    #########################
    # opensearch index (refresh)
    #########################
    print(f"deleting opensearch index... {vectorIndexName}") 
    
    try: # create index
        response = os_client.indices.delete(
            index_name
        )
        print('opensearch index was deleted:', response)
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)                
        #raise Exception ("Not able to create the index")        
    return 
   
# websocket
connection_url = os.environ.get('connection_url')
client = boto3.client('apigatewaymanagementapi', endpoint_url=connection_url)
print('connection_url: ', connection_url)

HUMAN_PROMPT = "\n\nHuman:"
AI_PROMPT = "\n\nAssistant:"

map_chain = dict() 

# Multi-LLM
def get_chat():
    global selected_chat
    
    if multi_region == 'enable':
        length_of_models = len(multi_region_models)
        profile = multi_region_models[selected_chat]
    else:
        length_of_models = len(LLM_for_chat)
        profile = LLM_for_chat[selected_chat]
        
    bedrock_region =  profile['bedrock_region']
    modelId = profile['model_id']
    maxOutputTokens = 4096
    print(f'LLM: {selected_chat}, bedrock_region: {bedrock_region}, modelId: {modelId}')
                          
    # bedrock   
    boto3_bedrock = boto3.client(
        service_name='bedrock-runtime',
        region_name=bedrock_region,
        config=Config(
            retries = {
                'max_attempts': 30
            }
        )
    )
    parameters = {
        "max_tokens":maxOutputTokens,     
        "temperature":0.1,
        "top_k":250,
        "top_p":0.9,
        "stop_sequences": [HUMAN_PROMPT]
    }
    # print('parameters: ', parameters)

    chat = ChatBedrock(   # new chat model
        model_id=modelId,
        client=boto3_bedrock, 
        model_kwargs=parameters,
    )    
    
    selected_chat = selected_chat + 1
    if selected_chat == length_of_models:
        selected_chat = 0
    
    return chat

def get_multi_region_chat(models, selected):
    profile = models[selected]
    bedrock_region =  profile['bedrock_region']
    modelId = profile['model_id']
    maxOutputTokens = 4096
    print(f'selected_chat: {selected}, bedrock_region: {bedrock_region}, modelId: {modelId}')
                          
    # bedrock   
    boto3_bedrock = boto3.client(
        service_name='bedrock-runtime',
        region_name=bedrock_region,
        config=Config(
            retries = {
                'max_attempts': 30
            }
        )
    )
    parameters = {
        "max_tokens":maxOutputTokens,     
        "temperature":0.1,
        "top_k":250,
        "top_p":0.9,
        "stop_sequences": [HUMAN_PROMPT]
    }
    # print('parameters: ', parameters)

    chat = ChatBedrock(   # new chat model
        model_id=modelId,
        client=boto3_bedrock, 
        model_kwargs=parameters,
    )    
    
    return chat

def get_multimodal():
    global selected_multimodal
    print('LLM_for_chat: ', LLM_for_chat)
    print('selected_multimodal: ', selected_multimodal)
        
    profile = LLM_for_multimodal[selected_multimodal]
    bedrock_region =  profile['bedrock_region']
    modelId = profile['model_id']
    maxOutputTokens = 4096
    print(f'LLM: {selected_multimodal}, bedrock_region: {bedrock_region}, modelId: {modelId}')
                          
    # bedrock   
    boto3_bedrock = boto3.client(
        service_name='bedrock-runtime',
        region_name=bedrock_region,
        config=Config(
            retries = {
                'max_attempts': 30
            }
        )
    )
    parameters = {
        "max_tokens":maxOutputTokens,     
        "temperature":0.1,
        "top_k":250,
        "top_p":0.9,
        "stop_sequences": [HUMAN_PROMPT]
    }
    # print('parameters: ', parameters)

    multimodal = ChatBedrock(   # new chat model
        model_id=modelId,
        client=boto3_bedrock, 
        model_kwargs=parameters,
    )    
    
    selected_multimodal = selected_multimodal + 1
    if selected_multimodal == len(LLM_for_multimodal):
        selected_multimodal = 0
    
    return multimodal

def get_embedding():
    global selected_embedding
    profile = LLM_embedding[selected_embedding]
    bedrock_region =  profile['bedrock_region']
    model_id = profile['model_id']
    print(f'selected_embedding: {selected_embedding}, bedrock_region: {bedrock_region}, model_id: {model_id}')
    
    # bedrock   
    boto3_bedrock = boto3.client(
        service_name='bedrock-runtime',
        region_name=bedrock_region, 
        config=Config(
            retries = {
                'max_attempts': 30
            }
        )
    )
    
    bedrock_embedding = BedrockEmbeddings(
        client=boto3_bedrock,
        region_name = bedrock_region,
        model_id = model_id
    )  
    
    selected_embedding = selected_embedding + 1
    if selected_embedding == len(LLM_embedding):
        selected_embedding = 0
    
    return bedrock_embedding
    
# load documents from s3 for pdf and txt
def load_document(file_type, s3_file_name):
    s3r = boto3.resource("s3")
    doc = s3r.Object(s3_bucket, s3_prefix+'/'+s3_file_name)
    
    if file_type == 'pdf':
        contents = doc.get()['Body'].read()
        reader = PyPDF2.PdfReader(BytesIO(contents))
        
        raw_text = []
        for page in reader.pages:
            raw_text.append(page.extract_text())
        contents = '\n'.join(raw_text)    
        
    elif file_type == 'txt':        
        contents = doc.get()['Body'].read().decode('utf-8')
        
    print('contents: ', contents)
    new_contents = str(contents).replace("\n"," ") 
    print('length: ', len(new_contents))

    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=100,
        separators=["\n\n", "\n", ".", " ", ""],
        length_function = len,
    ) 

    texts = text_splitter.split_text(new_contents) 
    print('texts[0]: ', texts[0])
    
    return texts

# load csv documents from s3
def load_csv_document(s3_file_name):
    s3r = boto3.resource("s3")
    doc = s3r.Object(s3_bucket, s3_prefix+'/'+s3_file_name)

    lines = doc.get()['Body'].read().decode('utf-8').split('\n')   # read csv per line
    print('lins: ', len(lines))
        
    columns = lines[0].split(',')  # get columns
    #columns = ["Category", "Information"]  
    #columns_to_metadata = ["type","Source"]
    print('columns: ', columns)
    
    docs = []
    n = 0
    for row in csv.DictReader(lines, delimiter=',',quotechar='"'):
        # print('row: ', row)
        #to_metadata = {col: row[col] for col in columns_to_metadata if col in row}
        values = {k: row[k] for k in columns if k in row}
        content = "\n".join(f"{k.strip()}: {v.strip()}" for k, v in values.items())
        doc = Document(
            page_content=content,
            metadata={
                'name': s3_file_name,
                'row': n+1,
            }
            #metadata=to_metadata
        )
        docs.append(doc)
        n = n+1
    print('docs[0]: ', docs[0])

    return docs

def get_summary(chat, docs):    
    text = ""
    for doc in docs:
        text = text + doc
    
    if isKorean(text)==True:
        system = (
            "다음의 <article> tag안의 문장을 요약해서 500자 이내로 설명하세오."
        )
    else: 
        system = (
            "Here is pieces of article, contained in <article> tags. Write a concise summary within 500 characters."
        )
    
    human = "<article>{text}</article>"
    
    prompt = ChatPromptTemplate.from_messages([("system", system), ("human", human)])
    print('prompt: ', prompt)
    
    chain = prompt | chat    
    try: 
        result = chain.invoke(
            {
                "text": text
            }
        )
        
        summary = result.content
        print('result of summarization: ', summary)
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)                    
        raise Exception ("Not able to request to LLM")
    
    return summary
    
def load_chatHistory(userId, allowTime, chat_memory):
    dynamodb_client = boto3.client('dynamodb')

    response = dynamodb_client.query(
        TableName=callLogTableName,
        KeyConditionExpression='user_id = :userId AND request_time > :allowTime',
        ExpressionAttributeValues={
            ':userId': {'S': userId},
            ':allowTime': {'S': allowTime}
        }
    )
    # print('query result: ', response['Items'])

    for item in response['Items']:
        text = item['body']['S']
        msg = item['msg']['S']
        type = item['type']['S']

        if type == 'text' and text and msg:
            memory_chain.chat_memory.add_user_message(text)
            memory_chain.chat_memory.add_ai_message(msg) 
                
def getAllowTime():
    d = datetime.datetime.now() - datetime.timedelta(days = 2)
    timeStr = str(d)[0:19]
    print('allow time: ',timeStr)

    return timeStr

def isKorean(text):
    # check korean
    pattern_hangul = re.compile('[\u3131-\u3163\uac00-\ud7a3]+')
    word_kor = pattern_hangul.search(str(text))
    # print('word_kor: ', word_kor)

    if word_kor and word_kor != 'None':
        print('Korean: ', word_kor)
        return True
    else:
        print('Not Korean: ', word_kor)
        return False

class GradeDocuments(BaseModel):
    """Binary score for relevance check on retrieved documents."""

    binary_score: str = Field(description="Documents are relevant to the question, 'yes' or 'no'")

def get_retrieval_grader(chat):
    system = """You are a grader assessing relevance of a retrieved document to a user question. \n 
    If the document contains keyword(s) or semantic meaning related to the question, grade it as relevant. \n
    Give a binary score 'yes' or 'no' score to indicate whether the document is relevant to the question."""

    grade_prompt = ChatPromptTemplate.from_messages(
        [
            ("system", system),
            ("human", "Retrieved document: \n\n {document} \n\n User question: {question}"),
        ]
    )
    
    structured_llm_grader = chat.with_structured_output(GradeDocuments)
    retrieval_grader = grade_prompt | structured_llm_grader
    return retrieval_grader

def grade_document_based_on_relevance(conn, question, doc, models, selected):     
    chat = get_multi_region_chat(models, selected)
    retrieval_grader = get_retrieval_grader(chat)
    
    #print(f"question: {question}, document:{doc.page_content}")    
    score = retrieval_grader.invoke({"question": question, "document": doc.page_content})
    #print(f"score: {score}")
    
    grade = score.binary_score    
    if grade == 'yes':
        print("---GRADE: DOCUMENT RELEVANT---")
        conn.send(doc)
    else:  # no
        print("---GRADE: DOCUMENT NOT RELEVANT---")
        conn.send(None)
    
    conn.close()
                                    
def grade_documents_using_parallel_processing(question, documents):
    global selected_chat
    
    filtered_docs = []    

    processes = []
    parent_connections = []
    
    for i, doc in enumerate(documents):
        #print(f"grading doc[{i}]: {doc.page_content}")        
        parent_conn, child_conn = Pipe()
        parent_connections.append(parent_conn)
            
        process = Process(target=grade_document_based_on_relevance, args=(child_conn, question, doc, multi_region_models, selected_chat))
        processes.append(process)

        selected_chat = selected_chat + 1
        if selected_chat == len(multi_region_models):
            selected_chat = 0
    for process in processes:
        process.start()
            
    for parent_conn in parent_connections:
        relevant_doc = parent_conn.recv()

        if relevant_doc is not None:
            filtered_docs.append(relevant_doc)

    for process in processes:
        process.join()
    
    #print('filtered_docs: ', filtered_docs)
    return filtered_docs

def get_ps_embedding():
    global selected_ps_embedding
    profile = priority_search_embedding[selected_ps_embedding]
    bedrock_region =  profile['bedrock_region']
    model_id = profile['model_id']
    print(f'selected_ps_embedding: {selected_ps_embedding}, bedrock_region: {bedrock_region}, model_id: {model_id}')
    
    # bedrock   
    boto3_bedrock = boto3.client(
        service_name='bedrock-runtime',
        region_name=bedrock_region, 
        config=Config(
            retries = {
                'max_attempts': 30
            }
        )
    )
    
    bedrock_ps_embedding = BedrockEmbeddings(
        client=boto3_bedrock,
        region_name = bedrock_region,
        model_id = model_id
    )  

    selected_ps_embedding = selected_ps_embedding + 1
    if selected_ps_embedding == len(priority_search_embedding):
        selected_ps_embedding = 0

    return bedrock_ps_embedding

def priority_search(query, relevant_docs, minSimilarity):
    excerpts = []
    for i, doc in enumerate(relevant_docs):
        #print('doc: ', doc)

        content = doc.page_content
        # print('content: ', content)

        excerpts.append(
            Document(
                page_content=content,
                metadata={
                    'name': doc.metadata['name'],
                    'url': doc.metadata['url'],
                    'from': doc.metadata['from'],
                    'order':i,
                    'score':0
                }
            )
        )
    #print('excerpts: ', excerpts)

    embeddings = get_ps_embedding()
    vectorstore_confidence = FAISS.from_documents(
        excerpts,  # documents
        embeddings  # embeddings
    )            
    rel_documents = vectorstore_confidence.similarity_search_with_score(
        query=query,
        k=len(relevant_docs)
    )

    docs = []
    for i, document in enumerate(rel_documents):
        print(f'## Document(priority_search) query: {query}, {i+1}: {document}')

        order = document[0].metadata['order']
        name = document[0].metadata['name']
        
        score = document[1]
        print(f"query: {query}, {order}: {name}, {score}")

        relevant_docs[order].metadata['score'] = int(score)

        if score < minSimilarity:
            docs.append(relevant_docs[order])    
    # print('selected docs: ', docs)

    return docs

def grade_documents(question, documents):
    print("###### grade_documents ######")
    
    filtered_docs = []
    print("start grading...")
    print("grade_state: ", grade_state)
    
    if grade_state == "LLM":
        if multi_region == 'enable':  # parallel processing        
            filtered_docs = grade_documents_using_parallel_processing(question, documents)

        else:
            # Score each doc    
            chat = get_chat()
            retrieval_grader = get_retrieval_grader(chat)
            for i, doc in enumerate(documents):
                # print('doc: ', doc)
                print_doc(i, doc)
                
                score = retrieval_grader.invoke({"question": question, "document": doc.page_content})
                # print("score: ", score)
                
                grade = score.binary_score
                # print("grade: ", grade)
                # Document relevant
                if grade.lower() == "yes":
                    print("---GRADE: DOCUMENT RELEVANT---")
                    filtered_docs.append(doc)
                # Document not relevant
                else:
                    print("---GRADE: DOCUMENT NOT RELEVANT---")
                    continue
    
    elif grade_state == "PRIORITY_SEARCH":
        filtered_docs = priority_search(question, documents, minDocSimilarity)
    else:  # OTHERS
        filtered_docs = documents
        
    # print('len(docments): ', len(filtered_docs))    
    return filtered_docs

def print_doc(i, doc):
    if len(doc.page_content)>=100:
        text = doc.page_content[:100]
    else:
        text = doc.page_content
            
    print(f"{i}: {text}, metadata:{doc.metadata}")

def query_using_RAG_context(connectionId, requestId, chat, context, revised_question):    
    isTyping(connectionId, requestId, "generating...")
    
    if isKorean(revised_question)==True:
        system = (
            """다음의 <context> tag안의 참고자료를 이용하여 상황에 맞는 구체적인 세부 정보를 충분히 제공합니다. Assistant의 이름은 서연이고, 모르는 질문을 받으면 솔직히 모른다고 말합니다.
            
            <context>
            {context}
            </context>"""
        )
    else: 
        system = (
            """Here is pieces of context, contained in <context> tags. Provide a concise answer to the question at the end. If you don't know the answer, just say that you don't know, don't try to make up an answer.
            
            <context>
            {context}
            </context>"""
        )
    
    human = "{input}"
    
    prompt = ChatPromptTemplate.from_messages([("system", system), ("human", human)])
    # print('prompt: ', prompt)
                   
    chain = prompt | chat
    
    try: 
        stream = chain.invoke(
            {
                "context": context,
                "input": revised_question,
            }
        )
        msg = readStreamMsg(connectionId, requestId, stream.content)    
        print('msg: ', msg)
        
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)        
            
        sendErrorMessage(connectionId, requestId, err_msg)    
        raise Exception ("Not able to request to LLM")

    return msg

def get_documents_from_opensearch(vectorstore_opensearch, query, top_k):
    try:
        result = vectorstore_opensearch.similarity_search_with_score(
            query = query,
            #k = top_k*2,  
            k = top_k,  
            search_type="script_scoring",
            pre_filter={"term": {"metadata.doc_level": "child"}}
        )    
        print('result: ', result)
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)        
           
    relevant_documents = []
    docList = []  # for duplication check
    for re in result:
        if 'parent_doc_id' in re[0].metadata:
            parent_doc_id = re[0].metadata['parent_doc_id']
            doc_level = re[0].metadata['doc_level']
            print(f"doc_level: {doc_level}, parent_doc_id: {parent_doc_id}")
                        
            if doc_level == 'child':
                if parent_doc_id in docList:
                    print('duplicated!')
                else:
                    relevant_documents.append(re)
                    docList.append(parent_doc_id)
                        
                    if len(relevant_documents)>=top_k:
                        break                                
    # print('lexical query result: ', json.dumps(response))
    
    #for i, doc in enumerate(relevant_documents):
    #    if len(doc[0].page_content)>=100:
    #        text = doc[0].page_content[:100]
    #    else:
    #        text = doc[0].page_content            
    #    print(f"--> vector search doc[{i}]: {text}, metadata:{doc[0].metadata}")

    return relevant_documents

os_client = OpenSearch(
    hosts = [{
        'host': opensearch_url.replace("https://", ""), 
        'port': 443
    }],
    http_compress = True,
    http_auth=(opensearch_account, opensearch_passwd),
    use_ssl = True,
    verify_certs = True,
    ssl_assert_hostname = False,
    ssl_show_warn = False,
)

def get_parent_content(parent_doc_id):
    response = os_client.get(
        index=index_name, 
        id = parent_doc_id
    )
    
    source = response['_source']                            
    # print('parent_doc: ', source['text'])   
    
    metadata = source['metadata']    
    #print('name: ', metadata['name'])   
    #print('url: ', metadata['url'])   
    #print('doc_level: ', metadata['doc_level']) 
    
    url = ""
    if "url" in metadata:
        url = metadata['url']
    
    return source['text'], metadata['name'], url

def lexical_search(query, top_k):
    # lexical search (keyword)
    min_match = 0
    
    query = {
        "query": {
            "bool": {
                "must": [
                    {
                        "match": {
                            "text": {
                                "query": query,
                                "minimum_should_match": f'{min_match}%',
                                "operator":  "or",
                            }
                        }
                    },
                ],
                "filter": [
                ]
            }
        }
    }

    response = os_client.search(
        body=query,
        index=index_name
    )
    print('lexical query result: ', json.dumps(response))
        
    docs = []
    for i, document in enumerate(response['hits']['hits']):
        if i>=top_k: 
            break
                    
        excerpt = document['_source']['text']
        
        name = document['_source']['metadata']['name']
        # print('name: ', name)

        page = ""
        if "page" in document['_source']['metadata']:
            page = document['_source']['metadata']['page']
        
        url = ""
        if "url" in document['_source']['metadata']:
            url = document['_source']['metadata']['url']            
        
        docs.append(
                Document(
                    page_content=excerpt,
                    metadata={
                        'name': name,
                        'url': url,
                        'page': page,
                        'from': 'lexical'
                    },
                )
            )
    
    for i, doc in enumerate(docs):
        #print('doc: ', doc)
        #print('doc content: ', doc.page_content)
        
        if len(doc.page_content)>=100:
            text = doc.page_content[:100]
        else:
            text = doc.page_content            
        print(f"--> lexical search doc[{i}]: {text}, metadata:{doc.metadata}")   
        
    return docs

def get_answer_using_opensearch(chat, text, connectionId, requestId):    
    global reference_docs
    
    msg = ""
    top_k = numberOfDocs
    relevant_docs = []
    
    bedrock_embedding = get_embedding()
       
    vectorstore_opensearch = OpenSearchVectorSearch(
        index_name = index_name,
        is_aoss = False,
        ef_search = 1024, # 512(default)
        m=48,
        #engine="faiss",  # default: nmslib
        embedding_function = bedrock_embedding,
        opensearch_url=opensearch_url,
        http_auth=(opensearch_account, opensearch_passwd), # http_auth=awsauth,
    )  
    
    isTyping(connectionId, requestId, "retrieving...")
    
    if enableParentDocumentRetrival == 'true': # parent/child chunking
        relevant_documents = get_documents_from_opensearch(vectorstore_opensearch, text, top_k)
                        
        for i, document in enumerate(relevant_documents):
            # print(f'## Document(opensearch-vector) {i+1}: {document}')
            
            parent_doc_id = document[0].metadata['parent_doc_id']
            doc_level = document[0].metadata['doc_level']
            #print(f"child: parent_doc_id: {parent_doc_id}, doc_level: {doc_level}")
            
            content, name, url = get_parent_content(parent_doc_id) # use pareant document
            #print(f"parent_doc_id: {parent_doc_id}, doc_level: {doc_level}, url: {url}, content: {content}")
            
            relevant_docs.append(
                Document(
                    page_content=content,
                    metadata={
                        'name': name,
                        'url': url,
                        'doc_level': doc_level,
                        'from': 'vector'
                    },
                )
            )
    else: 
        relevant_documents = vectorstore_opensearch.similarity_search_with_score(
            query = text,
            k = top_k
        )
        
        for i, document in enumerate(relevant_documents):
            # print(f'## Document(opensearch-vector) {i+1}: {document}')
            
            name = document[0].metadata['name']
            url = document[0].metadata['url']
            content = document[0].page_content
                   
            relevant_docs.append(
                Document(
                    page_content=content,
                    metadata={
                        'name': name,
                        'url': url,
                        'from': 'vector'
                    },
                )
            )
    
    if enableHybridSearch == 'true':
        relevant_docs += lexical_search(text, top_k)    
        
    isTyping(connectionId, requestId, "grading...")
    
    filtered_docs = grade_documents(text, relevant_docs) # grading
    
    filtered_docs = check_duplication(filtered_docs) # check duplication
            
    relevant_context = ""
    for i, document in enumerate(filtered_docs):
        # print(f"{i}: {document}")
        if document.page_content:
            content = document.page_content
            
        relevant_context = relevant_context + content + "\n\n"
        
    # print('relevant_context: ', relevant_context)

    msg = query_using_RAG_context(connectionId, requestId, chat, relevant_context, text)
    
    reference_docs += filtered_docs
           
    return msg

#########################################################

def check_duplication(docs):
    contentList = []
    length_original = len(docs)
    
    updated_docs = []
    print('length of relevant_docs:', len(docs))
    for doc in docs:            
        # print('excerpt: ', doc['metadata']['excerpt'])
            if doc.page_content in contentList:
                print('duplicated!')
                continue
            contentList.append(doc.page_content)
            updated_docs.append(doc)            
    length_updateed_docs = len(updated_docs)     
    
    if length_original == length_updateed_docs:
        print('no duplication')
    
    return updated_docs

def get_references(docs):
    reference = "\n\nFrom\n"
    
    cnt = 1
    nameList = []
    for i, doc in enumerate(docs):
        page = ""
        if "page" in doc.metadata:
            page = doc.metadata['page']
            #print('page: ', page)            
        url = ""
        if "url" in doc.metadata:
            url = doc.metadata['url']
            #print('url: ', url)                
        name = ""
        if "name" in doc.metadata:
            name = doc.metadata['name']
            #print('name: ', name)     
        pos = name.rfind('/')
        name = name[pos+1:]
        print(f"name: {name}")
           
        sourceType = ""
        if "from" in doc.metadata:
            sourceType = doc.metadata['from']
        #print('sourceType: ', sourceType)        
        
        #if len(doc.page_content)>=1000:
        #    excerpt = ""+doc.page_content[:1000]
        #else:
        #    excerpt = ""+doc.page_content
        excerpt = ""+doc.page_content
        # print('excerpt: ', excerpt)
        
        # for some of unusual case 
        #excerpt = excerpt.replace('"', '')        
        #excerpt = ''.join(c for c in excerpt if c not in '"')
        excerpt = re.sub('"', '', excerpt)
        # print('excerpt(quotation removed): ', excerpt)
        print('length: ', len(excerpt))
        
        if len(excerpt)<5000:
            if page:
                reference = reference + f"{cnt}. {page} page, <a href={url} target=_blank>{name}</a>, {sourceType}, <a href=\"#\" onClick=\"alert(`{excerpt}`)\">관련문서</a>\n"
            else:
                reference = reference + f"{cnt}. <a href={url} target=_blank>{name}</a>, {sourceType}, <a href=\"#\" onClick=\"alert(`{excerpt}`)\">관련문서</a>\n"
            cnt = cnt + 1
        else:            
            if name in nameList:
                print('duplicated!')
            else:
                #reference = reference + f"{cnt}. <a href={url} target=_blank>{name}</a>, {sourceType}\n"
                reference = reference + f"{cnt}. <a href={url} target=_blank>{name}</a>\n"
                nameList.append(name)                
                cnt = cnt+1
            
    return reference

def get_references_for_html(docs):
    reference = ""
    nameList = []
    cnt = 1
    for i, doc in enumerate(docs):
        print(f"reference {i}: doc")
        page = ""
        if "page" in doc.metadata:
            page = doc.metadata['page']
            #print('page: ', page)            
        url = ""
        if "url" in doc.metadata:
            url = doc.metadata['url']
            #print('url: ', url)                
        name = ""
        if "name" in doc.metadata:
            name = doc.metadata['name']
            #print('name: ', name)     
        pos = name.rfind('/')
        name = name[pos+1:]
        print(f"name: {name}")
           
        excerpt = ""+doc.page_content

        excerpt = re.sub('"', '', excerpt)
        print('length: ', len(excerpt))
        
        if name in nameList:
            print('duplicated!')
        else:
            reference = reference + f"{cnt}. <a href={url} target=_blank>{name}</a><br>"
            nameList.append(name)
            cnt = cnt+1
            
    return reference

def general_conversation(connectionId, requestId, chat, query):
    if isKorean(query)==True :
        system = (
            "다음의 Human과 Assistant의 친근한 이전 대화입니다. Assistant은 상황에 맞는 구체적인 세부 정보를 충분히 제공합니다. Assistant의 이름은 서연이고, 모르는 질문을 받으면 솔직히 모른다고 말합니다."
        )
    else: 
        system = (
            "Using the following conversation, answer friendly for the newest question. If you don't know the answer, just say that you don't know, don't try to make up an answer. You will be acting as a thoughtful advisor."
        )
    
    human = "{input}"
    
    prompt = ChatPromptTemplate.from_messages([("system", system), MessagesPlaceholder(variable_name="history"), ("human", human)])
    print('prompt: ', prompt)
    
    history = memory_chain.load_memory_variables({})["chat_history"]
    print('memory_chain: ', history)
                
    chain = prompt | chat    
    try: 
        isTyping(connectionId, requestId, "")  
        stream = chain.invoke(
            {
                "history": history,
                "input": query,
            }
        )
        msg = readStreamMsg(connectionId, requestId, stream.content)    
        
        usage = stream.response_metadata['usage']
        print('prompt_tokens: ', usage['prompt_tokens'])
        print('completion_tokens: ', usage['completion_tokens'])
        print('total_tokens: ', usage['total_tokens'])
        msg = stream.content

    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)        
            
        sendErrorMessage(connectionId, requestId, err_msg)    
        raise Exception ("Not able to request to LLM")
    
    return msg
  
def traslation(chat, text, input_language, output_language):
    system = (
        "You are a helpful assistant that translates {input_language} to {output_language} in <article> tags." 
        "Put it in <result> tags."
    )
    human = "<article>{text}</article>"
    
    prompt = ChatPromptTemplate.from_messages([("system", system), ("human", human)])
    # print('prompt: ', prompt)
    
    chain = prompt | chat    
    try: 
        result = chain.invoke(
            {
                "input_language": input_language,
                "output_language": output_language,
                "text": text,
            }
        )
        
        msg = result.content
        # print('translated text: ', msg)
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)                    
        raise Exception ("Not able to request to LLM")

    return msg[msg.find('<result>')+8:len(msg)-9] # remove <result> tag

def revise_question(connectionId, requestId, chat, query):    
    global history_length, token_counter_history    
    history_length = token_counter_history = 0
        
    isTyping(connectionId, requestId, "revising question...")
    
    if isKorean(query)==True :      
        system = (
            ""
        )  
        human = (
            "이전 대화를 참조하여, 다음의 <question>의 뜻을 명확히 하는 새로운 질문을 한국어로 생성하세요." 
            "새로운 질문은 원래 질문의 중요한 단어를 반드시 포함합니다. 결과는 <result> tag를 붙여주세요."
        
            "<question>"
            "{question}"
            "</question>"
        )    
    else: 
        system = (
            ""
        )
        human = (
            "Rephrase the follow up <question> to be a standalone question. Put it in <result> tags."
            "<question>"
            "{question}"
            "</question>"
        )
            
    prompt = ChatPromptTemplate.from_messages([("system", system), MessagesPlaceholder(variable_name="history"), ("human", human)])
    print('prompt: ', prompt)
    
    history = memory_chain.load_memory_variables({})["chat_history"]
    print('memory_chain: ', history)
                
    chain = prompt | chat    
    try: 
        result = chain.invoke(
            {
                "history": history,
                "question": query,
            }
        )
        generated_question = result.content
        
        revised_question = generated_question[generated_question.find('<result>')+8:len(generated_question)-9] # remove <result> tag                   
        print('revised_question: ', revised_question)
        
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)        
            
        sendErrorMessage(connectionId, requestId, err_msg)    
        raise Exception ("Not able to request to LLM")

    if debugMessageMode == 'true':  
        chat_history = ""
        for dialogue_turn in history:
            #print('type: ', dialogue_turn.type)
            #print('content: ', dialogue_turn.content)
            
            dialog = f"{dialogue_turn.type}: {dialogue_turn.content}\n"            
            chat_history = chat_history + dialog
                
        history_length = len(chat_history)
        print('chat_history length: ', history_length)
        
        token_counter_history = 0
        if chat_history:
            token_counter_history = chat.get_num_tokens(chat_history)
            print('token_size of history: ', token_counter_history)
            
        sendDebugMessage(connectionId, requestId, f"새로운 질문: {revised_question}\n * 대화이력({str(history_length)}자, {token_counter_history} Tokens)을 활용하였습니다.")
            
    return revised_question    
    # return revised_question.replace("\n"," ")

def isTyping(connectionId, requestId, msg):    
    if not msg:
        msg = "typing a message..."
    msg_proceeding = {
        'request_id': requestId,
        'msg': msg,
        'status': 'istyping'
    }
    #print('result: ', json.dumps(result))
    sendMessage(connectionId, msg_proceeding)
            
def readStreamMsg(connectionId, requestId, stream):
    msg = ""
    if stream:
        for event in stream:
            #print('event: ', event)
            msg = msg + event
            
            result = {
                'request_id': requestId,
                'msg': msg,
                'status': 'proceeding'
            }
            #print('result: ', json.dumps(result))
            sendMessage(connectionId, result)
    # print('msg: ', msg)
    return msg
    
def sendMessage(id, body):
    # print('sendMessage size: ', len(body))
    try:
        client.post_to_connection(
            ConnectionId=id, 
            Data=json.dumps(body)
        )
    except Exception:
        err_msg = traceback.format_exc()
        print('err_msg: ', err_msg)
        # raise Exception ("Not able to send a message")

def sendResultMessage(connectionId, requestId, msg):    
    result = {
        'request_id': requestId,
        'msg': msg,
        'status': 'completed'
    }
    #print('debug: ', json.dumps(debugMsg))
    sendMessage(connectionId, result)
    
def sendDebugMessage(connectionId, requestId, msg):
    debugMsg = {
        'request_id': requestId,
        'msg': msg,
        'status': 'debug'
    }
    #print('debug: ', json.dumps(debugMsg))
    sendMessage(connectionId, debugMsg)    
        
def sendErrorMessage(connectionId, requestId, msg):
    errorMsg = {
        'msg': msg,
        'request_id': requestId,
        'status': 'error'
    }
    print('error: ', json.dumps(errorMsg))
    sendMessage(connectionId, errorMsg)    

def load_chat_history(userId, allowTime):
    dynamodb_client = boto3.client('dynamodb')

    response = dynamodb_client.query(
        TableName=callLogTableName,
        KeyConditionExpression='user_id = :userId AND request_time > :allowTime',
        ExpressionAttributeValues={
            ':userId': {'S': userId},
            ':allowTime': {'S': allowTime}
        }
    )
    # print('query result: ', response['Items'])

    for item in response['Items']:
        text = item['body']['S']
        msg = item['msg']['S']
        type = item['type']['S']

        if type == 'text':
            memory_chain.chat_memory.add_user_message(text)
            memory_chain.chat_memory.add_ai_message(msg)     

def translate_text(chat, text):
    system = (
        "You are a helpful assistant that translates {input_language} to {output_language} in <article> tags." 
        "Put it in <result> tags."
    )
    human = "<article>{text}</article>"
    
    prompt = ChatPromptTemplate.from_messages([("system", system), ("human", human)])
    print('prompt: ', prompt)
    
    if isKorean(text)==False :
        input_language = "English"
        output_language = "Korean"
    else:
        input_language = "Korean"
        output_language = "English"
                        
    chain = prompt | chat    
    try: 
        result = chain.invoke(
            {
                "input_language": input_language,
                "output_language": output_language,
                "text": text,
            }
        )
        msg = result.content
        print('translated text: ', msg)
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)                    
        raise Exception ("Not able to request to LLM")

    return msg[msg.find('<result>')+8:len(msg)-9] # remove <result> tag

def check_grammer(chat, text):
    if isKorean(text)==True:
        system = (
            "다음의 <article> tag안의 문장의 오류를 찾아서 설명하고, 오류가 수정된 문장을 답변 마지막에 추가하여 주세요."
        )
    else: 
        system = (
            "Here is pieces of article, contained in <article> tags." 
            "Find the error in the sentence and explain it, and add the corrected sentence at the end of your answer."
        )
        
    human = "<article>{text}</article>"
    
    prompt = ChatPromptTemplate.from_messages([("system", system), ("human", human)])
    print('prompt: ', prompt)
    
    chain = prompt | chat    
    try: 
        result = chain.invoke(
            {
                "text": text
            }
        )
        
        msg = result.content
        print('result of grammer correction: ', msg)
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)                    
        raise Exception ("Not able to request to LLM")
    
    return msg

def use_multimodal(img_base64, query):    
    multimodal = get_multimodal()
    
    if query == "":
        query = "그림에 대해 상세히 설명해줘."
    
    messages = [
        SystemMessage(content="답변은 500자 이내의 한국어로 설명해주세요."),
        HumanMessage(
            content=[
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{img_base64}", 
                    },
                },
                {
                    "type": "text", "text": query
                },
            ]
        )
    ]
    
    try: 
        result = multimodal.invoke(messages)
        
        summary = result.content
        print('result of code summarization: ', summary)
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)                    
        raise Exception ("Not able to request to LLM")
    
    return summary

def extract_text(chat, img_base64):    
    query = "텍스트를 추출해서 utf8로 변환하세요. <result> tag를 붙여주세요."
    
    messages = [
        HumanMessage(
            content=[
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{img_base64}", 
                    },
                },
                {
                    "type": "text", "text": query
                },
            ]
        )
    ]
    
    try: 
        result = chat.invoke(messages)
        
        extracted_text = result.content
        print('result of text extraction from an image: ', extracted_text)
    except Exception:
        err_msg = traceback.format_exc()
        print('error message: ', err_msg)                    
        raise Exception ("Not able to request to LLM")
    
    return extracted_text

####################### LangGraph #######################
# Chat Agent Executor
#########################################################

@tool
def get_current_time(format: str=f"%Y-%m-%d %H:%M:%S")->str:
    """Returns the current date and time in the specified format"""
    print("###### get_current_time ######")
    # f"%Y-%m-%d %H:%M:%S"
    
    format = format.replace('\'','')
    timestr = datetime.datetime.now(timezone('Asia/Seoul')).strftime(format)
    # print('timestr:', timestr)
    
    return timestr

def get_lambda_client(region):
    return boto3.client(
        service_name='lambda',
        region_name=region
    )

@tool 
def get_book_list(keyword: str) -> str:
    """
    Search book list by keyword and then return book list
    keyword: search keyword
    return: book list
    """
    print("###### get_book_list ######")
    
    keyword = keyword.replace('\'','')

    answer = ""
    url = f"https://search.kyobobook.co.kr/search?keyword={keyword}&gbCode=TOT&target=total"
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        prod_info = soup.find_all("a", attrs={"class": "prod_info"})
        
        if len(prod_info):
            answer = "추천 도서는 아래와 같습니다.\n"
            
        for prod in prod_info[:5]:
            title = prod.text.strip().replace("\n", "")       
            link = prod.get("href")
            answer = answer + f"{title}, URL: {link}\n\n"
    
    return answer

@tool
def get_weather_info(city: str) -> str:
    """
    retrieve weather information by city name and then return weather statement.
    city: the name of city to retrieve
    return: weather statement
    """    
    print("###### get_weather_info ######")
    
    city = city.replace('\n','')
    city = city.replace('\'','')
    city = city.replace('\"','')
                
    chat = get_chat()
    if isKorean(city):
        place = traslation(chat, city, "Korean", "English")
        print('city (translated): ', place)
    else:
        place = city
        city = traslation(chat, city, "English", "Korean")
        print('city (translated): ', city)
        
    print('place: ', place)
    
    weather_str: str = f"{city}에 대한 날씨 정보가 없습니다."
    if weather_api_key: 
        apiKey = weather_api_key
        lang = 'en' 
        units = 'metric' 
        api = f"https://api.openweathermap.org/data/2.5/weather?q={place}&APPID={apiKey}&lang={lang}&units={units}"
        # print('api: ', api)
                
        try:
            result = requests.get(api)
            result = json.loads(result.text)
            print('result: ', result)
        
            if 'weather' in result:
                overall = result['weather'][0]['main']
                current_temp = result['main']['temp']
                min_temp = result['main']['temp_min']
                max_temp = result['main']['temp_max']
                humidity = result['main']['humidity']
                wind_speed = result['wind']['speed']
                cloud = result['clouds']['all']
                
                weather_str = f"{city}의 현재 날씨의 특징은 {overall}이며, 현재 온도는 {current_temp}도 이고, 최저온도는 {min_temp}도, 최고 온도는 {max_temp}도 입니다. 현재 습도는 {humidity}% 이고, 바람은 초당 {wind_speed} 미터 입니다. 구름은 {cloud}% 입니다."
                #weather_str = f"Today, the overall of {city} is {overall}, current temperature is {current_temp} degree, min temperature is {min_temp} degree, highest temperature is {max_temp} degree. huminity is {humidity}%, wind status is {wind_speed} meter per second. the amount of cloud is {cloud}%."            
        except Exception:
            err_msg = traceback.format_exc()
            print('error message: ', err_msg)                    
            # raise Exception ("Not able to request to LLM")    
        
    print('weather_str: ', weather_str)                            
    return weather_str

@tool
def search_by_tavily(keyword: str) -> str:
    """
    Search general information by keyword and then return the result as a string.
    keyword: search keyword which is greater than the minimum of 4 characters
    return: the information of keyword
    """    
    print("###### search_by_tavily ######")
    
    global reference_docs, selected_tavily
    
    docs = []
    if selected_tavily != -1:
        selected_tavily = selected_tavily + 1
        if selected_tavily == len(tavily_api_key):
            selected_tavily = 0

        try:
            tavily_client = TavilyClient(api_key=tavily_api_key[selected_tavily])
            response = tavily_client.search(keyword, max_results=3)
            # print('tavily response: ', response)
            
            print(f"--> tavily search result: {keyword}")
            for i, r in enumerate(response["results"]):
                content = r.get("content")
                print(f"{i}: {content}")

                name = r.get("title")
                if name is None:
                    name = 'WWW'
                    
                url = ""
                if "url" in r:
                    url = r.get("url")
            
                docs.append(
                    Document(
                        page_content=content,
                        metadata={
                            'name': name,
                            'url': url,
                            'from': 'tavily'
                        },
                    )
                )   
        except Exception as e:
            print('Exception: ', e)
        
        filtered_docs = grade_documents(keyword, docs)
        
        # duplication checker
        filtered_docs = check_duplication(filtered_docs)
        
    relevant_context = ""
    for i, document in enumerate(filtered_docs):
        print(f"{i}: {document}")
        if document.page_content:
            content = document.page_content
            
        relevant_context = relevant_context + content + "\n\n"        
    print('relevant_context: ', relevant_context)
        
    reference_docs += filtered_docs
        
    return relevant_context

@tool    
def search_by_opensearch(keyword: str) -> str:
    """
    Search information of company by keyword and then return the result as a string.
    keyword: search keyword
    return: the technical information of keyword
    """    
    global reference_docs
    
    print('keyword: ', keyword)
    keyword = keyword.replace('\'','')
    keyword = keyword.replace('|','')
    keyword = keyword.replace('\n','')
    print('modified keyword: ', keyword)
    
    bedrock_embedding = get_embedding()
        
    vectorstore_opensearch = OpenSearchVectorSearch(
        index_name = index_name,
        is_aoss = False,
        ef_search = 1024, # 512(default)
        m=48,
        #engine="faiss",  # default: nmslib
        embedding_function = bedrock_embedding,
        opensearch_url=opensearch_url,
        http_auth=(opensearch_account, opensearch_passwd), # http_auth=awsauth,
    ) 
    
    top_k = numberOfDocs
    relevant_docs = [] 
    if enableParentDocumentRetrival == 'true': # parent/child chunking
        relevant_documents = get_documents_from_opensearch(vectorstore_opensearch, keyword, top_k)
                        
        for i, document in enumerate(relevant_documents):
            #print(f'## Document(opensearch-vector) {i+1}: {document}')
            
            parent_doc_id = document[0].metadata['parent_doc_id']
            doc_level = document[0].metadata['doc_level']
            #print(f"child: parent_doc_id: {parent_doc_id}, doc_level: {doc_level}")
            
            excerpt, name, url = get_parent_content(parent_doc_id) # use pareant document
            #print(f"parent_doc_id: {parent_doc_id}, doc_level: {doc_level}, url: {url}, content: {excerpt}")
            
            relevant_docs.append(
                Document(
                    page_content=excerpt,
                    metadata={
                        'name': name,
                        'url': url,
                        'doc_level': doc_level,
                        'from': 'vector'
                    },
                )
            )
    else: 
        relevant_documents = vectorstore_opensearch.similarity_search_with_score(
            query = keyword,
            k = top_k
        )

        for i, document in enumerate(relevant_documents):
            #print(f'## Document(opensearch-vector) {i+1}: {document}')
            
            excerpt = document[0].page_content
            
            url = ""
            if "url" in document[0].metadata:
                url = document[0].metadata['url']
                
            name = document[0].metadata['name']
            
            relevant_docs.append(
                Document(
                    page_content=excerpt,
                    metadata={
                        'name': name,
                        'url': url,
                        'from': 'vector'
                    },
                )
            )
    
    if enableHybridSearch == 'true':
        relevant_docs += lexical_search(keyword, top_k)
    
    print('doc length: ', len(relevant_docs))
                
    filtered_docs = grade_documents(keyword, relevant_docs)  # grading
    
    filtered_docs = check_duplication(filtered_docs) # check duplication
    
    for i, doc in enumerate(filtered_docs):
        if len(doc.page_content)>=100:
            text = doc.page_content[:100]
        else:
            text = doc.page_content
            
        print(f"filtered doc[{i}]: {text}, metadata:{doc.metadata}")
       
    relevant_context = "" 
    for doc in filtered_docs:
        content = doc.page_content
        
        relevant_context = relevant_context + f"{content}\n\n"
        
    reference_docs += filtered_docs
        
    return relevant_context

def update_state_message(msg:str, config):
    print(msg)
    print('config: ', config)
    
    requestId = config.get("configurable", {}).get("requestId", "")
    connection = config.get("configurable", {}).get("connectionId", "")
    
    isTyping(connection, requestId, msg)
    
def run_agent_executor(connectionId, requestId, query):
    chatModel = get_chat() 
    tools = [get_current_time, get_book_list, get_weather_info, search_by_tavily, search_by_opensearch]

    model = chatModel.bind_tools(tools)

    class State(TypedDict):
        # messages: Annotated[Sequence[BaseMessage], operator.add]
        messages: Annotated[list, add_messages]
                
    tool_node = ToolNode(tools)

    def should_continue(state: State) -> Literal["continue", "end"]:
        print("###### should_continue ######")
        messages = state["messages"]    
        # print('(should_continue) messages: ', messages)
        
        last_message = messages[-1]
                
        if not last_message.tool_calls:
            next = "end"
        else:           
            next = "continue"     
        
        print(f"should_continue response: {next}")
        return next

    def call_model(state: State, config):
        print("###### call_model ######")
        
        update_state_message("thinking...", config)
        
        # print('state: ', state["messages"])
        
        if isKorean(state["messages"][0].content)==True:
            system = (
                "당신의 이름은 서연이고, 질문에 친근한 방식으로 대답하도록 설계된 대화형 AI입니다."
                "상황에 맞는 구체적인 세부 정보를 충분히 제공합니다."
                "모르는 질문을 받으면 솔직히 모른다고 말합니다."
                "최종 답변에는 조사한 내용을 반드시 포함합니다."
            )
        else: 
            system = (            
                "You are a conversational AI designed to answer in a friendly way to a question."
                "If you don't know the answer, just say that you don't know, don't try to make up an answer."
                "You will be acting as a thoughtful advisor."    
            )
            
        prompt = ChatPromptTemplate.from_messages(
            [
                ("system", system),
                MessagesPlaceholder(variable_name="messages"),
            ]
        )
        chain = prompt | model
            
        response = chain.invoke(state["messages"])
        print('call_model response: ', response.tool_calls)
        
        # state messag
        if response.tool_calls:
            toolinfo = response.tool_calls[-1]            
            if toolinfo['type'] == 'tool_call':
                print('tool name: ', toolinfo['name'])                    
                update_state_message(f"calling... {toolinfo['name']}", config)
        
        return {"messages": [response]}

    def buildChatAgent():
        workflow = StateGraph(State)

        workflow.add_node("agent", call_model)
        workflow.add_node("action", tool_node)
        workflow.add_edge(START, "agent")
        workflow.add_conditional_edges(
            "agent",
            should_continue,
            {
                "continue": "action",
                "end": END,
            },
        )
        workflow.add_edge("action", "agent")

        return workflow.compile()

    app = buildChatAgent()
        
    isTyping(connectionId, requestId, "")
    
    inputs = [HumanMessage(content=query)]
    config = {
        "recursion_limit": 50,
        "requestId": requestId,
        "connectionId": connectionId
    }
    
    message = ""
    for event in app.stream({"messages": inputs}, config, stream_mode="values"):   
        # print('event: ', event)
        
        message = event["messages"][-1]
        # print('message: ', message)

    msg = readStreamMsg(connectionId, requestId, message.content)
    
    # print('reference_docs: ', reference_docs)
    return msg

####################### LangGraph #######################
# Ocean Agent
#########################################################

class State(TypedDict):
    subject_company: str
    rating_date: str    
    planning_steps: List[str]
    sub_queries: List[List[str]]
    relevant_contexts : list[str]
    drafts : List[str]
    revised_drafts: Annotated[list, operator.add]  # for reflection
    references: List[str]
    
def markdown_to_html(body, reference):
    body = body + f"\n\n### 참고자료\n\n\n"
    
    html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    
    <md-block>
    </md-block>
    <script type="module" src="https://md-block.verou.me/md-block.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/github-markdown-css/5.2.0/github-markdown-light.css" integrity="sha512-n5zPz6LZB0QV1eraRj4OOxRbsV7a12eAGfFcrJ4bBFxxAwwYDp542z5M0w24tKPEhKk2QzjjIpR5hpOjJtGGoA==" crossorigin="anonymous" referrerpolicy="no-referrer"/>
</head>
<body>
    <div class="markdown-body">
        <md-block>{body}
        </md-block>
    </div>
    {reference}
</body>
</html>"""        
    return html

def get_documents_from_opensearch_for_subject_company(vectorstore_opensearch, query, top_k, subject_company):
    print(f"query: {query}, subject_company: {subject_company}")
    
    boolean_filter = {
        "bool": {
            "filter":[
                {"match" : {"metadata.subject_company":subject_company}},
                {"term" : {"metadata.doc_level":"child"}}
            ]
        }
    }          
    result = vectorstore_opensearch.similarity_search_with_score(
        query = query,
        k = top_k,
        search_type="script_scoring",
        pre_filter = boolean_filter
    )    
    print('result: ', result)
                
    relevant_documents = []
    docList = []
    for i, re in enumerate(result):
        #print(f"result[{i}] metadata: {re[0].metadata}")
        #print(f"result[{i}] page_content: {re[0].page_content}")
        
        parent_doc_id = doc_level = meta_subject_company = meta_rating_date = ""
        if "parent_doc_id" in re[0].metadata:
            parent_doc_id = re[0].metadata['parent_doc_id']
        if "doc_level" in re[0].metadata:
            doc_level = re[0].metadata['doc_level']
        if "subject_company" in re[0].metadata:
            meta_subject_company = re[0].metadata['subject_company']
        if "rating_date" in re[0].metadata:
            meta_rating_date = re[0].metadata['rating_date']
        print(f"--> (metadata) parent_doc_id: {parent_doc_id}, doc_level: {doc_level}, subject_company: {meta_subject_company}, rating_date: {meta_rating_date}")
             
        if parent_doc_id:
            if parent_doc_id in docList:
                print('duplicated!')
            else:
                relevant_documents.append(re)
                docList.append(parent_doc_id)
                        
                if len(relevant_documents)>=top_k:
                    break
    # print('lexical query result: ', json.dumps(response))
    
    #for i, doc in enumerate(relevant_documents):
        #print('doc: ', doc[0])
        #print('doc content: ', doc[0].page_content)
        
    #    if len(doc[0].page_content)>=100:
    #        text = doc[0].page_content[:100]
    #    else:
    #        text = doc[0].page_content            
    #    print(f"--> vector search doc[{i}]: {text}, metadata:{doc[0].metadata}")        

    return relevant_documents
        
def retrieve(query: str, subject_company: str):
    print(f'###### retrieve: {query} ######')
    
    top_k = numberOfDocs
    docs = []
    
    bedrock_embedding = get_embedding()
       
    vectorstore_opensearch = OpenSearchVectorSearch(
        index_name = index_name,
        is_aoss = False,
        ef_search = 1024, # 512(default)
        m=48,
        #engine="faiss",  # default: nmslib
        embedding_function = bedrock_embedding,
        opensearch_url=opensearch_url,
        http_auth=(opensearch_account, opensearch_passwd), # http_auth=awsauth,
    )  
    
    if enableParentDocumentRetrival == 'true': # parent/child chunking
        relevant_documents = get_documents_from_opensearch_for_subject_company(vectorstore_opensearch, query, top_k, subject_company)
                        
        for i, document in enumerate(relevant_documents):
            # print(f'## Document(opensearch-vector) {i+1}: {document}')
            
            parent_doc_id = document[0].metadata['parent_doc_id']
            doc_level = document[0].metadata['doc_level']
            #print(f"child: parent_doc_id: {parent_doc_id}, doc_level: {doc_level}")
            
            excerpt, name, url = get_parent_content(parent_doc_id) # use pareant document
            #print(f"parent_doc_id: {parent_doc_id}, doc_level: {doc_level}, url: {url}, content: {excerpt}")
            
            docs.append(
                Document(
                    page_content=excerpt,
                    metadata={
                        'name': name,
                        'url': url,
                        'doc_level': doc_level,
                        'from': 'vector'
                    },
                )
            )
    else: 
        boolean_filter = {
            "bool": {
                "filter":[
                    {"match" : {"metadata.subject_company":subject_company}}
                ]
            }
        }
        relevant_documents = vectorstore_opensearch.similarity_search_with_score(
            query = query,
            k = top_k,  
            search_type="script_scoring",
            pre_filter=boolean_filter
        )
        # print('result: ', result)
    
        for i, document in enumerate(relevant_documents):
            # print(f'## Document(opensearch-vector) {i+1}: {document}')
            
            name = document[0].metadata['name']
            url = document[0].metadata['url']
            content = document[0].page_content
                   
            docs.append(
                Document(
                    page_content=content,
                    metadata={
                        'name': name,
                        'url': url,
                        'from': 'vector'
                    },
                )
            )
    
    filtered_docs = grade_documents(query, docs) # grading

    filtered_docs = check_duplication(filtered_docs) # check duplication
            
    return filtered_docs

def parallel_retrieve(conn, q, subject_company):
    docs = retrieve(q, subject_company)                
    print(f"---> q: {q}, docs: {docs}")
                        
    conn.send(docs)    
    conn.close()

def retrieve_for_parallel_processing(sub_queries, subject_company):
    processes = []
    parent_connections = []
        
    relevant_docs = []
    for idx, q in enumerate(sub_queries):
        parent_conn, child_conn = Pipe()
        parent_connections.append(parent_conn)
            
        process = Process(target=parallel_retrieve, args=(child_conn, q, subject_company))
        processes.append(process)
            
    for process in processes:
        process.start()
                
    for parent_conn in parent_connections:
        docs = parent_conn.recv()
        
        relevant_docs += docs

    for process in processes:
        process.join()
          
    return relevant_docs 

def plan_node(state: State, config):
    print('###### plan_node ######')
    subject_company = state["subject_company"]    
    
    update_state_message('planning...', config)
    
    planning_steps = [
        "1. 회사 소개",
        "2. 주요 영업 활동",
        "3. 재무 현황",
        "4. 선대 현황",
        "5. 종합 평가"
    ]
    
    sub_queries = [
        [
            "establish", 
            "location", 
            "management", 
            "affiliated"
        ],
        [
            "cargo", 
            "route", 
            "owned/chartered", 
            "strategy"
        ],
        [
            "financial performance", 
            "route", 
            "financial risk",
            "payment"
        ],
        [
            "fleet"
        ],
        [
            "rating", #"infospectrum level"
            "assessment" # overall assessment"
        ]        
    ]
        
    return {
        "subject_company": subject_company,
        "planning_steps": planning_steps,
        "sub_queries": sub_queries
    }

def retrieve_node(state: State, config):
    print('###### retrieve_node ######')
    subject_company = state["subject_company"]    
    planning_steps = state["planning_steps"]
    print(f"subject_company: {subject_company}, planning_steps: {planning_steps}")
    
    update_state_message('retrieving...', config)
    
    relevant_contexts = []        
    references = []
    sub_queries = state["sub_queries"]
    
    for i, step in enumerate(planning_steps):
        update_state_message(f"retrieving... (step: {i+1}/{len(planning_steps)})", config)
        
        print(f"{i}: {step}")
        
        contents = ""    
        if multi_region == 'enable': 
            relevant_docs = retrieve_for_parallel_processing(sub_queries[i], subject_company)
            for doc in relevant_docs:
                contents += doc.page_content
            
            references += relevant_docs
        else:
                
            for q in sub_queries[i]:
                docs = retrieve(q, subject_company)
                
                print(f"---> q: {sub_queries[i]}, docs: {docs}")                
                for doc in docs:            
                    contents += doc.page_content
                
                references += docs
                                
        relevant_contexts.append(contents)
        
    return {
        "subject_company": subject_company,
        "planning_steps": planning_steps,
        "relevant_contexts": relevant_contexts,
        "references": references
    }

def generate_node(state: State, config):        
    print('###### generate_node ######')
    
    update_state_message('generating...', config)
    
    write_template = (
        "당신은 기업에 대한 보고서를 작성하는 훌륭한 글쓰기 도우미입니다."
        "아래와 같이 원본 보고서 지시사항과 계획한 보고서 단계를 제공하겠습니다."
        "또한 제가 이미 작성한 텍스트를 제공합니다."
        
        "보고서 지시사항:"
        "<instruction>"
        "{instruction}"
        "</instruction>"
        
        "보고서 단계:"
        "<plan>"
        "{plan}"
        "</plan>"
        
        "이미 작성한 텍스트:"
        "<text>"
        "{text}"
        "</text>"
        
        "참고 문서"
        "<context>"        
        "{context}"
        "</context>"        
        
        "보고서 지시 사항, 보고서 단계, 이미 작성된 텍스트, 참고 문서를 참조하여 다음 단계을 계속 작성합니다."
        "기업에 대한 구체적인 정보는 받드시 참고 문서를 이용해 작성하고, 모르는 부분은 포함하지 않습니다."
        
        "다음 단계:"
        "<step>"
        "{STEP}"
        "</step>"
                
        "보고서의 내용이 끊어지지 않고 잘 이해되도록 하나의 문단을 충분히 길게 작성합니다."
        "필요하다면 앞에 작은 부제를 추가할 수 있습니다."
        "이미 작성된 텍스트를 반복하지 말고 작성한 문단만 출력하세요."                
        "Markdown 포맷으로 서식을 작성하세요."
        "최종 결과에 <result> tag를 붙여주세요."
    )
    
    write_prompt = ChatPromptTemplate.from_messages([
        ("human", write_template)
    ])
    # print('prompt: ', prompt)
    
    instruction = f"{state['subject_company']} 회사에 대해 소개해 주세요."
    planning_steps = state["planning_steps"]
    text = ""
    drafts = []
    
    for i, step in enumerate(planning_steps):
        update_state_message(f"executing... (step: {i+1}/{len(planning_steps)})", config)
        
        print(f"{i}: {step}")
        context = state["relevant_contexts"][i]
        
        chat = get_chat()                       
        write_chain = write_prompt | chat            
        try: 
            result = write_chain.invoke({
                "instruction": instruction,
                "plan": planning_steps,
                "text": text,
                "context": context,
                "STEP": step
            })

            output = result.content
            draft = output[output.find('<result>')+8:len(output)-9] # remove <result> tag    
            print('draft: ', draft)
                
            if draft.find('#')!=-1 and draft.find('#')!=0:
                draft = draft[draft.find('#'):]
                    
            print(f"--> step:{step}")
            print(f"--> {draft}")

            text += draft + '\n\n'
            drafts.append(draft)
                
        except Exception:
            err_msg = traceback.format_exc()
            print('error message: ', err_msg)                        
            raise Exception ("Not able to request to LLM")

    return {
        "drafts": drafts
    }

#def should_continue(state: State):
#    print('###### continue_to_generate ######')
#    print('state (continue_to_generate): ', state)
    
#    current_step = state["current_step"]
#    planning_steps = state["planning_steps"]
#    print('current_step: ', current_step)
    
#    if current_step >= len(planning_steps):
#        return "end"
#    return "continue"
    
def buildOceanWorkflow():
    workflow = StateGraph(State)
    workflow.add_node("plan", plan_node)    
    workflow.add_node("retrieve", retrieve_node)        
    workflow.add_node("generate", generate_node)
    
    # Set entry point
    # workflow.set_entry_point("plan")
    workflow.add_edge(START, "plan")
    workflow.add_edge("plan", "retrieve")    
    workflow.add_edge("retrieve", "generate")
    workflow.add_edge("generate", END)
    
    return workflow.compile()

def get_final_answer(drafts, subject_company, references):
    final_doc = ""
    for i, draft in enumerate(drafts):
        print(f"{i}: {draft}")
        final_doc += draft

    # markdown file
    markdown_key = 'markdown/'+f"{subject_company}.md"
    # print('markdown_key: ', markdown_key)
        
    markdown_body = f"# {subject_company}\n\n"+final_doc
                
    s3_client = boto3.client('s3')  
    response = s3_client.put_object(
        Bucket=s3_bucket,
        Key=markdown_key,
        ContentType='text/markdown',
        Body=markdown_body.encode('utf-8')
    )
    # print('response: ', response)
        
    markdown_url = f"{path}{parse.quote(markdown_key)}"
    print('markdown_url: ', markdown_url)
        
    # html file
    html_key = 'markdown/'+f"{subject_company}.html"
        
    html_reference = ""
    print('references: ', references)
    if references:
        html_reference = get_references_for_html(references)
        
    html_body = markdown_to_html(markdown_body, html_reference)
    print('html_body: ', html_body)
        
    s3_client = boto3.client('s3')  
    response = s3_client.put_object(
        Bucket=s3_bucket,
        Key=html_key,
        ContentType='text/html',
        Body=html_body
    )
    # print('response: ', response)
        
    html_url = f"{path}{parse.quote(html_key)}"
    print('html_url: ', html_url)
    
    final_answer = final_doc+f"\n<a href={html_url} target=_blank>[미리보기 링크]</a>\n<a href={markdown_url} download=\"{subject_company}.md\">[다운로드 링크]</a>"    
    
    if references:
        reference = get_references(references)

    return final_answer+reference
            
def run_agent_ocean(connectionId, requestId, query):
    subject_company = query
    
    isTyping(connectionId, requestId, "")
    app = buildOceanWorkflow()
        
    # Run the workflow
    inputs = {
        "subject_company": subject_company
    }
    config = {
        "recursion_limit": 50,
        "requestId": requestId,
        "connectionId": connectionId
    }

    output = app.invoke(inputs, config)   
    # print('output: ', output)
    
    final_answer = get_final_answer(output['drafts'], subject_company, output['references'])
    
    return final_answer

####################### LangGraph #######################
# Ocean Agent Reflection
#########################################################

# Workflow - Reflection
class ReflectionState(TypedDict):
    draft : str
    subject_company: str
    reflection : List[str]
    search_queries : List[str]
    revised_draft: str
    revision_number: int    
        
class Reflection(BaseModel):
    missing: str = Field(description="Critique of what is missing.")
    advisable: str = Field(description="Critique of what is helpful for better writing")
    superfluous: str = Field(description="Critique of what is superfluous")

class Research(BaseModel):
    """Provide reflection and then follow up with search queries to improve the writing."""

    reflection: Reflection = Field(description="Your reflection on the initial writing.")
    search_queries: list[str] = Field(
        description="1-3 search queries for researching improvements to address the critique of your current writing."
    )

class ReflectionKor(BaseModel):
    missing: str = Field(description="작성된 글에 있어야하는데 빠진 내용이나 단점")
    advisable: str = Field(description="더 좋은 글이 되기 위해 추가하여야 할 내용")
    superfluous: str = Field(description="글의 길이나 스타일에 대한 비평")

class ResearchKor(BaseModel):
    """글쓰기를 개선하기 위한 검색 쿼리를 제공합니다."""

    reflection: ReflectionKor = Field(description="작성된 글에 대한 평가")
    search_queries: list[str] = Field(
        description="현재 글과 관련된 3개 이내의 검색어"
    )

def reflect_node(state: ReflectionState, config):
    print("###### reflect ######")
    draft = state['draft']
    print('draft: ', draft)
    subject_company = state['subject_company']
    print('subject_company: ', subject_company)
    
    idx = config.get("configurable", {}).get("idx")
    update_state_message(f"reflecting... (search_queries-{idx})", config)
    
    reflection = []
    search_queries = []
    search_queries_verified = []
    for attempt in range(5):
        chat = get_chat()
        
        if isKorean(draft):
            structured_llm = chat.with_structured_output(ResearchKor, include_raw=True)
        else:
            structured_llm = chat.with_structured_output(Research, include_raw=True)
            
        info = structured_llm.invoke(draft)
        print(f'attempt: {attempt}, info: {info}')
                
        if not info['parsed'] == None:
            parsed_info = info['parsed']
            # print('reflection: ', parsed_info.reflection)
            reflection = [parsed_info.reflection.missing, parsed_info.reflection.advisable]
            search_queries = parsed_info.search_queries
                
            print('reflection: ', parsed_info.reflection)            
            print('search_queries: ', search_queries)     
        
            #if isKorean(draft):
            #    translated_search = []
            #    for q in search_queries:
            #        chat = get_chat()
            #        if isKorean(q):
            #            search = traslation(chat, q, "Korean", "English")
            #        else:
            #            continue
            #            # search = traslation(chat, q, "English", "Korean")
            #        translated_search.append(search)
                        
            #    print('translated_search: ', translated_search)
            #    search_queries += translated_search

            #print('search_queries (mixed): ', search_queries)
            #search_queries_verified = search_queries
            #break
                        
            if isKorean(draft):
                for q in search_queries:
                    chat = get_chat()
                    if isKorean(q):
                        search = traslation(chat, q, "Korean", "English")
                        search_queries_verified.append(search)
                    else:
                        search_queries_verified.append(q)
                        
            print('search_queries (verified): ', search_queries_verified)
            break
        
    revision_number = state["revision_number"] if state.get("revision_number") is not None else 1
    return {
        "reflection": reflection,
        #"search_queries": search_queries,
        "search_queries": search_queries_verified,
        "revision_number": revision_number + 1,
        "subject_company": subject_company
    }

def revise_draft(state: ReflectionState, config):   
    print("###### revise_draft ######")
        
    draft = state['draft']
    search_queries = state['search_queries']
    reflection = state['reflection']
    print('draft: ', draft)
    print('search_queries: ', search_queries)
    print('reflection: ', reflection)
    
    idx = config.get("configurable", {}).get("idx")
    update_state_message(f"reflecting... (retrieve-{idx})", config)
        
    if isKorean(draft):
        revise_template = (
            "당신은 장문 작성에 능숙한 유능한 글쓰기 도우미입니다."                
            "draft을 critique과 information 사용하여 수정하십시오."
            "최종 결과는 한국어로 작성하고 <result> tag를 붙여주세요."
                            
            "<draft>"
            "{draft}"
            "</draft>"
                            
            "<critique>"
            "{reflection}"
            "</critique>"

            "<information>"
            "{content}"
            "</information>"
        )
    else:    
        revise_template = (
            "You are an excellent writing assistant." 
            "Revise this draft using the critique and additional information."
            # "Provide the final answer using Korean with <result> tag."
            "Provide the final answer with <result> tag."
                            
            "<draft>"
            "{draft}"
            "</draft>"
                        
            "<critique>"
            "{reflection}"
            "</critique>"

            "<information>"
            "{content}"
            "</information>"
        )
                    
    revise_prompt = ChatPromptTemplate([
        ('human', revise_template)
    ])
                              
    filtered_docs = []    
        
    # RAG - OpenSearch
    subject_company = state['subject_company']
    print('subject_company: ', subject_company)
    
    for q in search_queries:
        filtered_docs += retrieve(q, subject_company)
        print(f'q: {q}, filtered_docs: {filtered_docs}')
        
    # web search
    #for q in search_queries:
    #    docs = tavily_search(q, 4)
    #    print(f'q: {q}, WEB: {docs}')
        
    #    if len(docs):
    #        filtered_docs += grade_documents(q, docs)
    
    print('filtered_docs: ', filtered_docs)
              
    content = []   
    if len(filtered_docs):
        for d in filtered_docs:
            content.append(d.page_content)        
    # print('content: ', content)
    
    update_state_message(f"reflecting... (generate-{idx})", config)

    chat = get_chat()
    reflect = revise_prompt | chat
           
    res = reflect.invoke(
        {
            "draft": draft,
            "reflection": reflection,
            "content": content
        }
    )
    output = res.content
    # print('output: ', output)
        
    revised_draft = output[output.find('<result>')+8:len(output)-9]
    # print('revised_draft: ', revised_draft) 
            
    if revised_draft.find('#')!=-1 and revised_draft.find('#')!=0:
        revised_draft = revised_draft[revised_draft.find('#'):]

    print('--> draft: ', draft)
    print('--> reflection: ', reflection)
    print('--> revised_draft: ', revised_draft)
        
    revision_number = state["revision_number"] if state.get("revision_number") is not None else 1
        
    return {
        "revised_draft": revised_draft,
        "revision_number": revision_number
    }
    
MAX_REVISIONS = 1
def should_continue(state: ReflectionState, config):
    print("###### should_continue ######")
    max_revisions = config.get("configurable", {}).get("max_revisions", MAX_REVISIONS)
    revision_number = state["revision_number"]
    print(f"revision_number: {revision_number}, max_revisions: {max_revisions}")
            
    if revision_number > max_revisions:
        return "end"
    return "continue"

def buildReflection():
    workflow = StateGraph(ReflectionState)

    # Add nodes
    workflow.add_node("reflect_node", reflect_node)
    workflow.add_node("revise_draft", revise_draft)

    # Set entry point
    workflow.set_entry_point("reflect_node")
        
    workflow.add_conditional_edges(
        "revise_draft", 
        should_continue, 
        {
            "end": END, 
            "continue": "reflect_node"
        }
    )

    # Add edges
    workflow.add_edge("reflect_node", "revise_draft")
        
    return workflow.compile()

def reflect_draft(conn, reflection_app, config, idx, draft, subject_company):     
    inputs = {
        "draft": draft,
        "subject_company": subject_company,
        "revision_number": 0
    }        
    output = reflection_app.invoke(inputs, config)
        
    result = {
        "revised_draft": output['revised_draft'],
        "idx": idx
    }
            
    conn.send(result)    
    conn.close()

def reflect_drafts_using_parallel_processing(drafts, config, subject_company):
    revised_drafts = drafts
        
    processes = []
    parent_connections = []
        
    reflection_app = buildReflection()
    
    requestId = config.get("configurable", {}).get("requestId", "")
    print('requestId: ', requestId)
    connectionId = config.get("configurable", {}).get("connectionId", "")
    print('connectionId: ', connectionId)
                
    for idx, draft in enumerate(drafts):
        parent_conn, child_conn = Pipe()
        parent_connections.append(parent_conn)
            
        app_config = {
            "recursion_limit": 50,
            "max_revisions": MAX_REVISIONS,
            "requestId":requestId,
            "connectionId": connectionId,
            "idx": idx
        }
        process = Process(target=reflect_draft, args=(child_conn, reflection_app, app_config, idx, draft, subject_company))
        processes.append(process)
            
    for process in processes:
        process.start()
                
    for parent_conn in parent_connections:
        result = parent_conn.recv()

        if result is not None:
            print('result: ', result)
            revised_drafts[result['idx']] = result['revised_draft']

    for process in processes:
        process.join()
          
    return revised_drafts 

def revise_answers(state: State, config):
    print("###### revise_answers ######")
    drafts = state["drafts"]
    print('drafts: ', drafts)
    subject_company = state['subject_company']
    print('subject_company: ', subject_company)
    
    update_state_message('revising...', config)
    
    # reflection
    if multi_region == 'enable':  # parallel processing
        revised_drafts = reflect_drafts_using_parallel_processing(drafts, config, subject_company)
    else:
        revised_drafts = []
        reflection_app = buildReflection()
                
        for idx, draft in enumerate(drafts):
            inputs = {
                "draft": draft,
                "subject_company": subject_company
            }                
            output = reflection_app.invoke(inputs, config)
            
            revised_drafts.append(output['revised_draft'])
                
    return {"revised_drafts": revised_drafts}
            
def buildPlanAndExecuteOceanWorkflow():
    workflow = StateGraph(State)
    
    # Add nodes
    workflow.add_node("plan", plan_node)
    workflow.add_node("retrieve", retrieve_node)        
    workflow.add_node("generate", generate_node)
    workflow.add_node("revise_answers", revise_answers)  # reflection
    
    # Set entry point
    workflow.set_entry_point("plan")    
    
    # Add edges
    workflow.add_edge("plan", "retrieve")
    workflow.add_edge("retrieve", "generate")
    workflow.add_edge("generate", "revise_answers")
    workflow.add_edge("revise_answers", END)
    
    return workflow.compile()
    
def run_agent_ocean_reflection(connectionId, requestId, query):
    subject_company = query
    
    isTyping(connectionId, requestId, "")

    app = buildPlanAndExecuteOceanWorkflow()
        
    # Run the workflow
    inputs = {
        "subject_company": subject_company
    }
    config = {
        "recursion_limit": 50,
        "requestId": requestId,
        "connectionId": connectionId
    }

    output = app.invoke(inputs, config)   
    # print('output: ', output)
    
    final_answer = get_final_answer(output['revised_drafts'], subject_company, output['references'])
    
    return final_answer
    
def getResponse(connectionId, jsonBody):
    print('jsonBody: ', jsonBody)
    
    userId  = jsonBody['user_id']
    print('userId: ', userId)
    requestId  = jsonBody['request_id']
    print('requestId: ', requestId)
    requestTime  = jsonBody['request_time']
    print('requestTime: ', requestTime)
    type  = jsonBody['type']
    print('type: ', type)
    body = jsonBody['body']
    print('body: ', body)
    convType = jsonBody['convType']
    print('convType: ', convType)
    
    global multi_region    
    if "multi_region" in jsonBody:
        multi_region = jsonBody['multi_region']
    print('multi_region: ', multi_region)
        
    print('initiate....')

    global map_chain, memory_chain
    
    # Multi-LLM
    if multi_region == 'enable':
        profile = multi_region_models[selected_chat]
    else:
        profile = LLM_for_chat[selected_chat]
    bedrock_region =  profile['bedrock_region']
    modelId = profile['model_id']
    # print(f'selected_chat: {selected_chat}, bedrock_region: {bedrock_region}, modelId: {modelId}')
    # print('profile: ', profile)
    
    global grade_state
    if "grade" in jsonBody:
        grade_state = jsonBody['grade']
    else:
        grade_state = 'NONE'
    print('grade_state: ', grade_state)
    
    global reference_docs, contentList
    reference_docs = []
    contentList = []
    
    chat = get_chat()    
    
    # create memory
    if userId in map_chain:  
        print('memory exist. reuse it!')
        memory_chain = map_chain[userId]
    else: 
        print('memory does not exist. create new one!')        
        memory_chain = ConversationBufferWindowMemory(memory_key="chat_history", output_key='answer', return_messages=True, k=5)
        map_chain[userId] = memory_chain

        allowTime = getAllowTime()
        load_chat_history(userId, allowTime)
    
    msg = ""
    reference = ""
    
    if type == 'text' and body[:11] == 'list models':
        bedrock_client = boto3.client(
            service_name='bedrock',
            region_name=bedrock_region,
        )
        modelInfo = bedrock_client.list_foundation_models()    
        print('models: ', modelInfo)

        msg = f"The list of models: \n"
        lists = modelInfo['modelSummaries']
        
        for model in lists:
            msg += f"{model['modelId']}\n"
        
        msg += f"current model: {modelId}"
        print('model lists: ', msg)    
        
    elif type == 'text' and body[:21] == 'refresh current index':
        # refresh index
        isTyping(connectionId, requestId, "")
        refresh_opensearch_index()
        msg = "The index was refreshed in OpenSearch."
        sendResultMessage(connectionId, requestId, msg)
        
    else:             
        if type == 'text':
            text = body
            print('query: ', text)

            querySize = len(text)
            textCount = len(text.split())
            print(f"query size: {querySize}, words: {textCount}")

            if text == 'clearMemory':
                memory_chain.clear()
                map_chain[userId] = memory_chain
                    
                print('initiate the chat memory!')
                msg  = "The chat memory was intialized in this session."
            else:            
                if convType == 'normal':      # normal
                    msg = general_conversation(connectionId, requestId, chat, text)                  
                
                elif convType == 'rag-opensearch':   # RAG - Vector
                    msg = get_answer_using_opensearch(chat, text, connectionId, requestId)
                    
                    if reference_docs:
                        reference = get_references(reference_docs)
                
                elif convType == 'rag-opensearch-chat':   # RAG - Vector
                    revised_question = revise_question(connectionId, requestId, chat, text)     
                    print('revised_question: ', revised_question)  
                    
                    msg = get_answer_using_opensearch(chat, revised_question, connectionId, requestId)
                    
                    if reference_docs:
                        reference = get_references(reference_docs)
                
                elif convType == 'agent-executor':
                    msg = run_agent_executor(connectionId, requestId, text)
                    
                    if reference_docs:
                        reference = get_references(reference_docs)
                
                elif convType == 'agent-executor-chat':
                    revised_question = revise_question(connectionId, requestId, chat, text)     
                    print('revised_question: ', revised_question)  
                    msg = run_agent_executor(connectionId, requestId, revised_question)
                    
                    if reference_docs:
                        reference = get_references(reference_docs)
                    
                elif convType == "agent-ocean":
                    msg = run_agent_ocean(connectionId, requestId, text)
                
                elif convType == "agent-ocean-reflection":
                    msg = run_agent_ocean_reflection(connectionId, requestId, text)
                                    
                memory_chain.chat_memory.add_user_message(text)
                memory_chain.chat_memory.add_ai_message(msg)
                
        elif type == 'document':
            isTyping(connectionId, requestId, "")
            
            object = body
            file_type = object[object.rfind('.')+1:len(object)]            
            print('file_type: ', file_type)
            
            if file_type == 'csv':
                if not convType == "bedrock-agent":
                    docs = load_csv_document(object)
                    contexts = []
                    for doc in docs:
                        contexts.append(doc.page_content)
                    print('contexts: ', contexts)
                
                    msg = get_summary(chat, contexts)

            elif file_type == 'pdf' or file_type == 'txt' or file_type == 'md' or file_type == 'pptx' or file_type == 'docx':
                texts = load_document(file_type, object)

                docs = []
                for i in range(len(texts)):
                    docs.append(
                        Document(
                            page_content=texts[i],
                            metadata={
                                'name': object,
                                # 'page':i+1,
                                'url': path+doc_prefix+parse.quote(object)
                            }
                        )
                    )
                print('docs[0]: ', docs[0])    
                print('docs size: ', len(docs))

                contexts = []
                for doc in docs:
                    contexts.append(doc.page_content)
                print('contexts: ', contexts)

                msg = get_summary(chat, contexts)
                                
            elif file_type == 'png' or file_type == 'jpeg' or file_type == 'jpg':
                print('multimodal: ', object)
                
                s3_client = boto3.client('s3') 
                    
                image_obj = s3_client.get_object(Bucket=s3_bucket, Key=s3_prefix+'/'+object)
                # print('image_obj: ', image_obj)
                
                image_content = image_obj['Body'].read()
                img = Image.open(BytesIO(image_content))
                
                width, height = img.size 
                print(f"width: {width}, height: {height}, size: {width*height}")
                
                isResized = False
                while(width*height > 5242880):                    
                    width = int(width/2)
                    height = int(height/2)
                    isResized = True
                    print(f"width: {width}, height: {height}, size: {width*height}")
                
                if isResized:
                    img = img.resize((width, height))
                
                buffer = BytesIO()
                img.save(buffer, format="PNG")
                img_base64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
                
                command = ""        
                if 'command' in jsonBody:
                    command  = jsonBody['command']
                    print('command: ', command)
                
                # verify the image
                msg = use_multimodal(img_base64, command)       
                
                # extract text from the image
                text = extract_text(chat, img_base64)
                extracted_text = text[text.find('<result>')+8:len(text)-9] # remove <result> tag
                print('extracted_text: ', extracted_text)
                if len(extracted_text)>10:
                    msg = msg + f"\n\n[추출된 Text]\n{extracted_text}\n"
                
                memory_chain.chat_memory.add_user_message(f"{object}에서 텍스트를 추출하세요.")
                memory_chain.chat_memory.add_ai_message(extracted_text)
                
            else:
                msg = "uploaded file: "+object
        
        sendResultMessage(connectionId, requestId, msg+reference)
        # print('msg+reference: ', msg+reference)    
                
        print('msg: ', msg)

        item = {
            'user_id': {'S':userId},
            'request_id': {'S':requestId},
            'request_time': {'S':requestTime},
            'type': {'S':type},
            'body': {'S':body},
            'msg': {'S':msg+reference}
        }
        client = boto3.client('dynamodb')
        try:
            resp =  client.put_item(TableName=callLogTableName, Item=item)
        except Exception:
            err_msg = traceback.format_exc()
            print('error message: ', err_msg)
            # raise Exception ("Not able to write into dynamodb")         
        #print('resp, ', resp)

    return msg, reference

def lambda_handler(event, context):
    # print('event: ', event)
    
    msg = ""
    if event['requestContext']: 
        connectionId = event['requestContext']['connectionId']        
        routeKey = event['requestContext']['routeKey']
        
        if routeKey == '$connect':
            print('connected!')
        elif routeKey == '$disconnect':
            print('disconnected!')
        else:
            body = event.get("body", "")
            #print("data[0:8]: ", body[0:8])
            if body[0:8] == "__ping__":
                # print("keep alive!")                
                sendMessage(connectionId, "__pong__")
            else:
                print('connectionId: ', connectionId)
                print('routeKey: ', routeKey)
        
                jsonBody = json.loads(body)
                print('request body: ', json.dumps(jsonBody))

                requestId  = jsonBody['request_id']
                
                start = int(time.time())  
                try:
                    msg, reference = getResponse(connectionId, jsonBody)

                    print('msg+reference: ', msg+reference)
                    
                    elapsed_time = int(time.time()) - start
                    print("total run time(sec): ", elapsed_time)
                                        
                except Exception:
                    err_msg = traceback.format_exc()
                    print('err_msg: ', err_msg)

                    sendErrorMessage(connectionId, requestId, err_msg)    
                    raise Exception ("Not able to send a message")

    return {
        'statusCode': 200
    }
