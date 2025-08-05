"""
Supabase Client for STRIX
"""
from supabase import create_client, Client
from typing import List, Dict, Any, Optional
import logging
from config import SUPABASE_URL, SUPABASE_SERVICE_KEY

logger = logging.getLogger(__name__)

class SupabaseClient:
    def __init__(self):
        self.client: Client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
        
    def create_tables(self):
        """Initialize database tables"""
        # Note: 실제로는 Supabase Dashboard에서 생성하는 것이 권장됨
        # 여기서는 참고용 SQL만 제공
        
        sql_statements = [
            """
            CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
            CREATE EXTENSION IF NOT EXISTS "vector";
            """,
            
            """
            CREATE TABLE IF NOT EXISTS documents (
                id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
                type VARCHAR(50) NOT NULL,
                source VARCHAR(255),
                title TEXT,
                organization VARCHAR(100),
                category VARCHAR(100),
                created_at TIMESTAMP DEFAULT NOW(),
                file_path TEXT,
                metadata JSONB
            );
            """,
            
            """
            CREATE TABLE IF NOT EXISTS chunks (
                id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
                document_id UUID REFERENCES documents(id) ON DELETE CASCADE,
                content TEXT NOT NULL,
                chunk_index INTEGER,
                start_char INTEGER,
                end_char INTEGER,
                metadata JSONB,
                created_at TIMESTAMP DEFAULT NOW()
            );
            """,
            
            """
            CREATE TABLE IF NOT EXISTS embeddings (
                id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
                chunk_id UUID REFERENCES chunks(id) ON DELETE CASCADE,
                embedding vector(1536),
                model VARCHAR(50) DEFAULT 'text-embedding-ada-002',
                created_at TIMESTAMP DEFAULT NOW()
            );
            """,
            
            """
            CREATE TABLE IF NOT EXISTS correlations (
                id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
                internal_doc_id UUID REFERENCES documents(id),
                external_doc_id UUID REFERENCES documents(id),
                score FLOAT,
                reasoning TEXT,
                verified BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMP DEFAULT NOW()
            );
            """,
            
            """
            CREATE TABLE IF NOT EXISTS search_logs (
                id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
                query TEXT,
                results JSONB,
                user_feedback JSONB,
                created_at TIMESTAMP DEFAULT NOW()
            );
            """,
            
            """
            CREATE TABLE IF NOT EXISTS keyword_learning (
                id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
                category VARCHAR(50),
                keyword VARCHAR(100),
                frequency INTEGER DEFAULT 1,
                relevance_score FLOAT,
                status VARCHAR(20) DEFAULT 'pending',
                created_at TIMESTAMP DEFAULT NOW()
            );
            """
        ]
        
        logger.info("SQL statements for table creation prepared")
        return sql_statements
    
    def insert_document(self, doc_data: Dict[str, Any]) -> str:
        """Insert a document and return its ID"""
        try:
            response = self.client.table('documents').insert(doc_data).execute()
            return response.data[0]['id']
        except Exception as e:
            logger.error(f"Error inserting document: {e}")
            raise
    
    def insert_chunks(self, chunks_data: List[Dict[str, Any]]):
        """Insert multiple chunks"""
        try:
            response = self.client.table('chunks').insert(chunks_data).execute()
            return response.data
        except Exception as e:
            logger.error(f"Error inserting chunks: {e}")
            raise
    
    def insert_embeddings(self, embeddings_data: List[Dict[str, Any]]):
        """Insert embeddings"""
        try:
            response = self.client.table('embeddings').insert(embeddings_data).execute()
            return response.data
        except Exception as e:
            logger.error(f"Error inserting embeddings: {e}")
            raise
    
    def search_similar_chunks(self, query_embedding: List[float], 
                            limit: int = 10,
                            filters: Optional[Dict] = None) -> List[Dict]:
        """Search for similar chunks using vector similarity"""
        try:
            # Prepare RPC parameters matching the function signature
            rpc_params = {
                'query_embedding': query_embedding,
                'match_threshold': 0.7,
                'match_count': limit
            }
            
            # Map filter keys to the expected parameter names
            if filters:
                if 'type' in filters:
                    rpc_params['filter_type'] = filters['type']
                if 'category' in filters:
                    rpc_params['filter_category'] = filters['category']
                if 'organization' in filters:
                    rpc_params['filter_organization'] = filters['organization']
            
            response = self.client.rpc('search_chunks', rpc_params).execute()
            return response.data
        except Exception as e:
            logger.error(f"Error searching chunks: {e}")
            return []
    
    def get_document_by_id(self, doc_id: str) -> Optional[Dict]:
        """Get document by ID"""
        try:
            response = self.client.table('documents').select("*").eq('id', doc_id).execute()
            return response.data[0] if response.data else None
        except Exception as e:
            logger.error(f"Error getting document: {e}")
            return None
    
    def insert_correlation(self, correlation_data: Dict[str, Any]):
        """Insert correlation between documents"""
        try:
            response = self.client.table('correlations').insert(correlation_data).execute()
            return response.data
        except Exception as e:
            logger.error(f"Error inserting correlation: {e}")
            raise
    
    def get_correlations(self, doc_id: str, doc_type: str = 'internal') -> List[Dict]:
        """Get correlations for a document"""
        try:
            if doc_type == 'internal':
                response = self.client.table('correlations')\
                    .select("*, external_doc:external_doc_id(title, category, created_at)")\
                    .eq('internal_doc_id', doc_id)\
                    .order('score', desc=True)\
                    .execute()
            else:
                response = self.client.table('correlations')\
                    .select("*, internal_doc:internal_doc_id(title, organization, created_at)")\
                    .eq('external_doc_id', doc_id)\
                    .order('score', desc=True)\
                    .execute()
            
            return response.data
        except Exception as e:
            logger.error(f"Error getting correlations: {e}")
            return []
    
    def log_search(self, query: str, results: Dict, metadata: Optional[Dict] = None):
        """Log search query and results"""
        try:
            log_data = {
                'query': query,
                'results': results
            }
            self.client.table('search_logs').insert(log_data).execute()
        except Exception as e:
            logger.error(f"Error logging search: {e}")
    
    def suggest_keywords(self, category: str, min_frequency: int = 5) -> List[str]:
        """Get keyword suggestions for a category"""
        try:
            response = self.client.table('keyword_learning')\
                .select("keyword, frequency, relevance_score")\
                .eq('category', category)\
                .eq('status', 'approved')\
                .gte('frequency', min_frequency)\
                .order('relevance_score', desc=True)\
                .limit(20)\
                .execute()
            
            return [item['keyword'] for item in response.data]
        except Exception as e:
            logger.error(f"Error getting keyword suggestions: {e}")
            return []