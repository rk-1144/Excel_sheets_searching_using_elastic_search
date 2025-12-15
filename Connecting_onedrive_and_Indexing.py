
import msal
import requests
import io
import pandas as pd
from elasticsearch import Elasticsearch, helpers
from datetime import datetime

# ============ CONFIGURATION ============
CLIENT_ID = ""  # ← PUT YOUR CLIENT_ID HERE
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.Read.All", "User.Read"]
ONEDRIVE_FOLDER = "Excel"  # ← Your OneDrive folder name
INDEX_NAME = 'excel_fields_data'  # ← Elasticsearch index name

# ============ ELASTICSEARCH CONNECTION ============
es = Elasticsearch(['http://localhost:9200'])

def authenticate_onedrive():
    """Authenticate with OneDrive using device code flow"""
    print("🔐 Authenticating with OneDrive...\n")
    
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    
    # Initiate device code flow
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception("Failed to create device flow")
    
    print(flow["message"])
    print("\n⏳ Waiting for authentication...\n")
    
    # Wait for user to authenticate
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise Exception(f"Authentication failed: {result}")
    
    print("✅ Successfully authenticated with OneDrive!\n")
    return result["access_token"]


def create_elasticsearch_index():
    """Create Elasticsearch index with mapping for your new Excel structure"""
    mapping = {
        "properties": {
            # Your new Excel columns (based on the image)
            "field_name": {"type": "text"},
            "description": {"type": "text"},
            "field_type": {"type": "keyword"},
            "format": {"type": "text"},
            "field_length": {"type": "text"},
            "default_value": {"type": "text"},
            "valid_values": {"type": "text"},
            "field_behaviour": {"type": "text"},
            "visibility_rules": {"type": "text"},
            "visibility_attributes": {"type": "text"},
            
            # Metadata
            "filename": {"type": "keyword"},
            "row_number": {"type": "integer"},
            "indexed_at": {"type": "date"}
        }
    }
    
    # Delete existing index
    if es.indices.exists(index=INDEX_NAME):
        es.indices.delete(index=INDEX_NAME)
        print(f"✓ Deleted existing index: {INDEX_NAME}")
    
    # Create new index
    es.indices.create(index=INDEX_NAME, mappings=mapping)
    print(f"✓ Created new index: {INDEX_NAME}\n")


def clean_value(value):
    """Clean cell values"""
    if pd.isna(value):
        return None
    return str(value).strip()


def index_excel_from_onedrive(access_token, folder_path):
    """
    Read Excel files from OneDrive and index them into Elasticsearch
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}:/children"
    
    print("="*70)
    print("FETCHING FILES FROM ONEDRIVE AND INDEXING TO ELASTICSEARCH")
    print("="*70 + "\n")
    
    # Get list of files in OneDrive folder
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"❌ Error accessing OneDrive folder: {response.status_code}")
        print(response.text)
        return
    
    data = response.json()
    items = data.get("value", [])
    
    # Filter Excel files
    excel_files = [item for item in items if item.get("name", "").lower().endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print(f"❌ No Excel files found in OneDrive folder: {folder_path}")
        return
    
    print(f"📁 Found {len(excel_files)} Excel file(s) in OneDrive\n")
    
    total_docs = 0
    
    for i, item in enumerate(excel_files, 1):
        filename = item["name"]
        file_id = item["id"]
        
        print(f"[{i}/{len(excel_files)}] 📄 Processing: {filename}")
        
        try:
            # Download file content
            content_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
            file_response = requests.get(content_url, headers=headers)
            
            if file_response.status_code != 200:
                print(f"   ✗ Failed to download file: {file_response.status_code}")
                continue
            
            # Read Excel file into DataFrame
            df = pd.read_excel(io.BytesIO(file_response.content))
            
            print(f"   → Found {len(df)} rows")
            
            # Prepare documents for Elasticsearch
            actions = []
            
            for idx, row in df.iterrows():
                # Map your Excel columns to Elasticsearch fields
                # Adjust column names based on your actual Excel headers
                doc = {
                    'field_name': clean_value(row.get('Field Name')),
                    'description': clean_value(row.get('Description')),
                    'field_type': clean_value(row.get('Field Type')),
                    'format': clean_value(row.get('Format')),
                    'field_length': clean_value(row.get('Field Length')),
                    'default_value': clean_value(row.get('Default Value')),
                    'valid_values': clean_value(row.get('Valid Values')),
                    'field_behaviour': clean_value(row.get('Field Behaviour')),
                    'visibility_rules': clean_value(row.get('Visibility Rules')),
                    'visibility_attributes': clean_value(row.get('Visibility Attributes')),
                    'filename': filename,
                    'row_number': idx + 2,  # +2 because Excel row 1 is header
                    'indexed_at': datetime.now().isoformat()
                }
                
                action = {
                    "_index": INDEX_NAME,
                    "_source": doc
                }
                actions.append(action)
            
            # Bulk index to Elasticsearch
            if actions:
                success, failed = helpers.bulk(es, actions, raise_on_error=False)
                print(f"   ✓ Indexed {success} documents")
                if failed:
                    print(f"   ✗ Failed: {len(failed)} documents")
                total_docs += success
            
        except Exception as e:
            print(f"   ✗ Error processing file: {str(e)}")
        
        print()
    
    print("="*70)
    print(f"✅ INDEXING COMPLETE!")
    print(f"Total documents indexed: {total_docs}")
    print("="*70)
    
    # Refresh index to make data searchable immediately
    print("\n🔄 Refreshing index...")
    es.indices.refresh(index=INDEX_NAME)
    
    return total_docs


def verify_and_show_samples():
    """Verify indexing and show sample data"""
    print("\n" + "="*70)
    print("VERIFICATION")
    print("="*70 + "\n")
    
    # Count documents
    count = es.count(index=INDEX_NAME)
    print(f"📊 Total documents in Elasticsearch: {count['count']}\n")
    
    # Show sample documents
    result = es.search(
        index=INDEX_NAME,
        query={"match_all": {}},
        size=3
    )
    
    print("📋 Sample documents:\n")
    for i, hit in enumerate(result['hits']['hits'], 1):
        doc = hit['_source']
        print(f"Document {i}:")
        print(f"  Field Name: {doc.get('field_name')}")
        print(f"  Description: {doc.get('description')}")
        print(f"  Field Type: {doc.get('field_type')}")
        print(f"  Format: {doc.get('format')}")
        print(f"  File: {doc.get('filename')}, Row: {doc.get('row_number')}")
        print()


def test_search():
    """Test search functionality"""
    print("="*70)
    print("TESTING SEARCH")
    print("="*70 + "\n")
    
    # Test search
    result = es.search(
        index=INDEX_NAME,
        query={
            "multi_match": {
                "query": "text",
                "fields": ["field_name", "description", "field_type"]
            }
        },
        size=5
    )
    
    print(f"🔍 Search for 'text': Found {result['hits']['total']['value']} results\n")
    for hit in result['hits']['hits']:
        doc = hit['_source']
        print(f"  • {doc.get('field_name')} - {doc.get('field_type')} ({doc.get('filename')})")


if __name__ == "__main__":
    print("\n" + "="*70)
    print("🚀 ONEDRIVE TO ELASTICSEARCH INDEXER")
    print("="*70 + "\n")
    
    # Check Elasticsearch connection
    print("Checking Elasticsearch connection...")
    try:
        if not es.ping():
            print("❌ Cannot connect to Elasticsearch!")
            print("Make sure Elasticsearch is running: docker ps")
            exit(1)
        print("✓ Connected to Elasticsearch\n")
    except Exception as e:
        print(f"❌ Connection error: {e}")
        exit(1)
    
    # Authenticate with OneDrive
    try:
        access_token = authenticate_onedrive()
    except Exception as e:
        print(f"❌ OneDrive authentication failed: {e}")
        exit(1)
    
    # Create Elasticsearch index
    create_elasticsearch_index()
    
    # Index files from OneDrive
    total_docs = index_excel_from_onedrive(access_token, ONEDRIVE_FOLDER)
    
    if total_docs > 0:
        # Verify and show samples
        verify_and_show_samples()
        
        # Test search
        test_search()
    
    print("\n✅ ALL DONE! Your OneDrive Excel data is now searchable in Elasticsearch")
    print("Next step: Create the search API and web interface\n")
