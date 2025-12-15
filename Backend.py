from flask import Flask, request, jsonify
from flask_cors import CORS
from elasticsearch import Elasticsearch
import os
import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__)
CORS(app)

# Elasticsearch configuration - Compatible with ES 8.x
es = Elasticsearch(
    ['http://localhost:9200'],
    verify_certs=False,
    ssl_show_warn=False,
    request_timeout=30
)

INDEX_NAME = 'excel_fields_data'

# Excel directory (keeping for future OneDrive integration)
EXCEL_DIR = os.path.join(os.path.dirname(__file__), 'excel-files')

# Hard-coded file list (as provided)
HARDCODED_FILES = [
    "Customer_Support_Fields.xlsx",
    "Employee_Management_Fields.xlsx",
    "Event_Registration_Fields.xlsx",
    "Inventory_Tracking_Fields.xlsx",
    "Invoice_Management_Fields.xlsx",
    "Order_Processing_Fields.xlsx",
    "portfolio_data.xlsx",
    "Product_Catalog_Fields.xlsx",
    "Project_Management_Fields.xlsx",
    "Survey_Response_Fields.xlsx",
    "User_Registration_Fields.xlsx"
]

@app.route('/api/excel-files', methods=['GET'])
def get_excel_files():
    """
    Return a hard-coded list of files for the frontend dropdown.
    Edit `HARDCODED_FILES` above to change what is shown.
    """
    try:
        print("\n📂 Returning hard-coded file list for frontend")
        files = [
            {"id": os.path.splitext(f)[0], "name": f}
            for f in HARDCODED_FILES
            if f and (f.endswith('.xlsx') or f.endswith('.xls'))
        ]
       # print(f"  Files: {[f['name'] for f in files]}")
        return {"files": files}
    except Exception as e:
        print(f"❌ Error building hard-coded file list: {e}")
        return {"error": str(e), "files": []}, 500

# @app.route('/api/excel-files', methods=['GET'])
# def get_excel_files():
#     """
#     Get unique list of files from Elasticsearch index.
#     If ES is unreachable or returns no buckets, fall back to reading local `excel-files` directory.
#     """
#     try:
#         print("\n📂 Fetching file list from Elasticsearch...")

#         # If ES isn't responding, fallback immediately to disk
#         try:
#             if not es.ping():
#                 print("❌ Elasticsearch is not responding (ping failed) - falling back to local files")
#                 raise RuntimeError("es.ping() failed")
#         except Exception as ping_exc:
#             print(f"   ping check error: {ping_exc}")
#             # fallthrough to local fallback below

#             # local fallback
#             local_files = [
#                 {"id": os.path.splitext(f)[0].lower(), "name": f}
#                 for f in os.listdir(EXCEL_DIR)
#                 if f.lower().endswith('.xlsx') or f.lower().endswith('.xls')
#             ]
#             print(f"ℹ️  Returning {len(local_files)} file(s) from local folder")
#             return {"files": local_files}

#         # Query ES for unique filenames
#         try:
#             result = es.search(
#                 index=INDEX_NAME,
#                 body={
#                     "size": 0,
#                     "aggs": {
#                         "unique_files": {
#                             "terms": {
#                                 "field": "filename.keyword",
#                                 "size": 100
#                             }
#                         }
#                     }
#                 }
#             )
#             # debug log the top-level keys to help troubleshoot
#             print("  ES response keys:", list(result.keys()))
#         except Exception as es_exc:
#             print(f"❌ Elasticsearch query failed: {es_exc}")
#             # fall back to local directory if ES search fails
#             local_files = [
#                 {"id": os.path.splitext(f)[0].lower(), "name": f}
#                 for f in os.listdir(EXCEL_DIR)
#                 if f.lower().endswith('.xlsx') or f.lower().endswith('.xls')
#             ]
#             print(f"ℹ️  Returning {len(local_files)} file(s) from local folder (ES failed)")
#             return {"files": local_files}

#         # Safely extract buckets
#         buckets = []
#         aggs = result.get('aggregations') or result.get('agg') or {}
#         if aggs and 'unique_files' in aggs and 'buckets' in aggs['unique_files']:
#             buckets = aggs['unique_files']['buckets']
#         else:
#             # No aggregation results - fallback to local files
#             print("⚠️  No 'unique_files' aggregation found in ES response. Falling back to local folder.")
#             local_files = [
#                 {"id": os.path.splitext(f)[0].lower(), "name": f}
#                 for f in os.listdir(EXCEL_DIR)
#                 if f.lower().endswith('.xlsx') or f.lower().endswith('.xls')
#             ]
#             print(f"ℹ️  Returning {len(local_files)} file(s) from local folder")
#             return {"files": local_files}

#         files = [
#             {
#                 "id": os.path.splitext(bucket.get('key', ''))[0].lower(),
#                 "name": bucket.get('key', '')
#             }
#             for bucket in buckets
#             if bucket.get('key')
#         ]

#         print(f"✅ Found {len(files)} file(s) from Elasticsearch:")
#         for f in files:
#             print(f"   - {f['name']}")

#         return {"files": files}

#     except Exception as e:
#         print(f"❌ Error fetching files: {e}")
#         import traceback
#         traceback.print_exc()
#         # Always return a consistent JSON structure so frontend won't break
#         # Try to return local files as last resort
#         try:
#             local_files = [
#                 {"id": os.path.splitext(f)[0].lower(), "name": f}
#                 for f in os.listdir(EXCEL_DIR)
#                 if f.lower().endswith('.xlsx') or f.lower().endswith('.xls')
#             ]
#         except Exception as disk_e:
#             print(f"  also failed to read local folder: {disk_e}")
#             local_files = []
#         return {"error": str(e), "files": local_files}, 500


@app.route('/api/search-excel', methods=['POST'])
def search_excel():
    """
    Search Excel data using Elasticsearch with partial matching
    Accepts: fileName, fieldName, fieldType, visibilityRules, visibilityAttributes
    """
    try:
        params = request.json
        
        # Get search parameters (all optional)
        filename = params.get('fileName', '').strip()
        field_name = params.get('fieldName', '').strip()
        field_type = params.get('fieldType', '').strip()
        visibility_rules = params.get('visibilityRules', '').strip()
        visibility_attributes = params.get('visibilityAttributes', '').strip()
        
        print(f"\n{'='*70}")
        print(f"🔍 SEARCH REQUEST")
        print('='*70)
        print(f"  File: '{filename}'")
        print(f"  Field Name: '{field_name}'")
        print(f"  Field Type: '{field_type}'")
        print(f"  Visibility Rules: '{visibility_rules}'")
        print(f"  Visibility Attributes: '{visibility_attributes}'")
        
        # Build Elasticsearch query
        must_conditions = []
        
        # Add file filter if specified
        if filename:
            if not filename.endswith('.xlsx'):
                filename = filename + '.xlsx'

            must_conditions.append({
                            "term": {
                            "filename": filename
        }
    })

        # Add field_name search with partial matching
        if field_name:
            must_conditions.append({
                "bool": {
                    "should": [
                        {"wildcard": {"field_name": f"*{field_name.lower()}*"}},
                        {"wildcard": {"field_name": f"*{field_name.upper()}*"}},
                        {"wildcard": {"field_name": f"*{field_name.title()}*"}},
                        {"match": {"field_name": field_name}}
                    ],
                    "minimum_should_match": 1
                }
            })
        
        # Add field_type search with partial matching
        if field_type:
            must_conditions.append({
                "bool": {
                    "should": [
                        {"wildcard": {"field_type": f"*{field_type.lower()}*"}},
                        {"wildcard": {"field_type": f"*{field_type.upper()}*"}},
                        {"wildcard": {"field_type": f"*{field_type.title()}*"}},
                        {"match": {"field_type": field_type}}
                    ],
                    "minimum_should_match": 1
                }
            })
        
        # Add visibility_rules search with partial matching
        if visibility_rules:
            must_conditions.append({
                "bool": {
                    "should": [
                        {"wildcard": {"visibility_rules": f"*{visibility_rules.lower()}*"}},
                        {"wildcard": {"visibility_rules": f"*{visibility_rules.upper()}*"}},
                        {"wildcard": {"visibility_rules": f"*{visibility_rules.title()}*"}},
                        {"match": {"visibility_rules": visibility_rules}}
                    ],
                    "minimum_should_match": 1
                }
            })
        
        # Add visibility_attributes search with partial matching
        if visibility_attributes:
            must_conditions.append({
                "bool": {
                    "should": [
                        {"wildcard": {"visibility_attributes": f"*{visibility_attributes.lower()}*"}},
                        {"wildcard": {"visibility_attributes": f"*{visibility_attributes.upper()}*"}},
                        {"wildcard": {"visibility_attributes": f"*{visibility_attributes.title()}*"}},
                        {"match": {"visibility_attributes": visibility_attributes}}
                    ],
                    "minimum_should_match": 1
                }
            })
        
        # If no conditions, return all documents
        if not must_conditions:
            search_query = {"match_all": {}}
            print("  ℹ️  No filters - returning all documents")
        else:
            search_query = {
                "bool": {
                    "must": must_conditions
                }
            }
        
        # Execute search
        result = es.search(
            index=INDEX_NAME,
            query=search_query,
            size=1000  # Adjust based on your needs
        )
        
        total = result['hits']['total']['value']
        print(f"\n✅ Found {total} result(s)")
        
        # Format results to match frontend expectations
        results = []
        for hit in result['hits']['hits']:
            doc = hit['_source']
            
            # Map Elasticsearch fields to frontend format (camelCase)
            result_item = {
                'fieldName': doc.get('field_name', ''),
                'description': doc.get('description', ''),
                'fieldType': doc.get('field_type', ''),
                'format': doc.get('format', ''),
                'fieldLength': doc.get('field_length', ''),
                'defaultValue': doc.get('default_value', ''),
                'validValues': doc.get('valid_values', ''),
                'fieldBehaviour': doc.get('field_behaviour', ''),
                'visibilityRules': doc.get('visibility_rules', ''),
                'visibilityAttributes': doc.get('visibility_attributes', ''),
                'sourceFile': doc.get('filename', ''),
                'rowNumber': doc.get('row_number', '')
            }
            
            results.append(result_item)
        
        print(f"✅ Returning {len(results)} results to frontend")
        
        # Print first result as sample
        if results:
            print(f"\n📄 Sample result:")
            print(f"   Field: {results[0]['fieldName']}")
            print(f"   Type: {results[0]['fieldType']}")
            print(f"   File: {results[0]['sourceFile']}")
        
        return {"results": results}
    
    except Exception as e:
        print(f"❌ Error in search: {e}")
        import traceback
        traceback.print_exc()
        return {"error": str(e), "results": []}, 500


@app.route('/api/health', methods=['GET'])
def health_check():
    """
    Check if Elasticsearch is connected and healthy
    """
    try:
        if es.ping():
            # Get index stats
            stats = es.count(index=INDEX_NAME)
            return {
                "status": "healthy",
                "elasticsearch": "connected",
                "index": INDEX_NAME,
                "document_count": stats['count']
            }
        else:
            return {
                "status": "unhealthy",
                "elasticsearch": "disconnected"
            }, 500
    except Exception as e:
        return {
            "status": "error",
            "message": str(e)
        }, 500


@app.route('/api/debug/field-types', methods=['GET'])
def get_field_types():
    """
    Debug endpoint to get all unique field types in the index
    """
    try:
        result = es.search(
            index=INDEX_NAME,
            body={
                "size": 0,
                "aggs": {
                    "unique_types": {
                        "terms": {
                            "field": "field_type.keyword",
                            "size": 100
                        }
                    }
                }
            }
        )
        
        buckets = result['aggregations']['unique_types']['buckets']
        field_types = [
            {
                "type": bucket['key'],
                "count": bucket['doc_count']
            }
            for bucket in buckets
        ]
        
        return {"fieldTypes": field_types}
    
    except Exception as e:
        return {"error": str(e)}, 500


@app.route('/api/debug/sample-doc', methods=['GET'])
def get_sample_doc():
    """
    Get a sample document to see the actual field structure
    """
    try:
        result = es.search(
            index=INDEX_NAME,
            query={"match_all": {}},
            size=1
        )
        
        if result['hits']['hits']:
            doc = result['hits']['hits'][0]['_source']
            return {"sample": doc}
        else:
            return {"error": "No documents found"}, 404
    
    except Exception as e:
        return {"error": str(e)}, 500


if __name__ == '__main__':
    print("\n" + "="*70)
    print("🚀 STARTING FLASK SERVER WITH ELASTICSEARCH")
    print("="*70)
    
    # Check Elasticsearch connection on startup
    print("\n📊 Checking Elasticsearch connection...")
    try:
        if es.ping():
            print("✅ Connected to Elasticsearch at http://localhost:9200")
            try:
                stats = es.count(index=INDEX_NAME)
                print(f"✅ Index '{INDEX_NAME}' has {stats['count']} documents")
                
                # Get a sample document to show field structure
                sample = es.search(index=INDEX_NAME, query={"match_all": {}}, size=1)
                if sample['hits']['hits']:
                    doc = sample['hits']['hits'][0]['_source']
                    print(f"\n📄 Sample document fields:")
                    for key in doc.keys():
                        print(f"   - {key}: {doc[key]}")
                
            except Exception as e:
                print(f"⚠️  Warning with index '{INDEX_NAME}': {e}")
        else:
            print("❌ Cannot connect to Elasticsearch!")
            print("   Make sure Elasticsearch is running in Docker")
            print("   Run: docker ps | grep elasticsearch")
    except Exception as e:
        print(f"❌ Elasticsearch connection error: {e}")
        print("   Make sure Docker container is running")
        import traceback
        traceback.print_exc()
    
    # Ensure Excel directory exists (for future OneDrive integration)
    if not os.path.exists(EXCEL_DIR):
        os.makedirs(EXCEL_DIR)
        print(f"\n📁 Created directory: {EXCEL_DIR}")
    
    print("\n" + "="*70)
    print("🌐 Server starting on http://localhost:3001")
    print("="*70)
    print("\nAvailable endpoints:")
    print("  GET  /api/excel-files          - List all files")
    print("  POST /api/search-excel         - Search records")
    print("  GET  /api/health               - Health check")
    print("  GET  /api/debug/field-types    - Debug field types")
    print("  GET  /api/debug/sample-doc     - See sample document")
    print("\n💡 Test in browser:")
    print("   http://localhost:3001/api/health")
    print("   http://localhost:3001/api/excel-files")
    print("   http://localhost:3001/api/debug/sample-doc")
    print("\n" + "="*70 + "\n")
    
    app.run(host='0.0.0.0', port=3001, debug=True)