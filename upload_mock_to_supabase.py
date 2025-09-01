"""
Upload Mock Data to Supabase Database
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

import json
import glob
from datetime import datetime
from pathlib import Path
from langchain_openai import OpenAIEmbeddings
from database.supabase_client import SupabaseClient
from config import CHUNK_SIZE, CHUNK_OVERLAP
from langchain.text_splitter import RecursiveCharacterTextSplitter
from dotenv import load_dotenv
import time

load_dotenv()

class MockDataUploader:
    def __init__(self):
        self.supabase = SupabaseClient()
        self.embeddings = OpenAIEmbeddings()
        self.text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=CHUNK_SIZE,
            chunk_overlap=CHUNK_OVERLAP
        )
        
    def clear_existing_data(self):
        """기존 데이터 삭제"""
        print("\n[1/5] Clearing existing data...")
        
        try:
            # embeddings 테이블 삭제 (FK 제약 때문에 먼저)
            self.supabase.client.table('embeddings').delete().neq('id', '00000000-0000-0000-0000-000000000000').execute()
            print("  - Cleared embeddings table")
            
            # chunks 테이블 삭제
            self.supabase.client.table('chunks').delete().neq('id', '00000000-0000-0000-0000-000000000000').execute()
            print("  - Cleared chunks table")
            
            # documents 테이블 삭제
            self.supabase.client.table('documents').delete().neq('id', '00000000-0000-0000-0000-000000000000').execute()
            print("  - Cleared documents table")
            
            print("  [OK] All existing data cleared")
            return True
            
        except Exception as e:
            print(f"  [ERROR] Failed to clear data: {str(e)}")
            return False
    
    def load_metadata(self):
        """메타데이터 파일 로드"""
        metadata_path = "./mock_data/metadata.json"
        
        if os.path.exists(metadata_path):
            with open(metadata_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}
    
    def process_internal_documents(self):
        """내부 문서 처리"""
        print("\n[2/5] Processing internal documents...")
        
        internal_dir = Path("./mock_data/internal")
        doc_count = 0
        
        for org_dir in internal_dir.iterdir():
            if org_dir.is_dir():
                organization = org_dir.name
                print(f"\n  Processing {organization}...")
                
                for txt_file in org_dir.glob("*.txt"):
                    try:
                        # 문서 읽기
                        with open(txt_file, 'r', encoding='utf-8') as f:
                            content = f.read()
                        
                        # 문서 정보 생성
                        doc_data = {
                            'type': 'internal',
                            'source': organization,
                            'title': txt_file.stem.replace('_', ' '),
                            'organization': organization,
                            'category': self._extract_category(txt_file.stem),
                            'file_path': str(txt_file),
                            'metadata': {
                                'file_type': 'TXT',
                                'loaded_at': datetime.now().isoformat()
                            }
                        }
                        
                        # 문서 저장
                        doc_result = self.supabase.client.table('documents').insert(doc_data).execute()
                        document_id = doc_result.data[0]['id']
                        
                        # 청킹
                        chunks = self.text_splitter.split_text(content)
                        
                        # 청크 및 임베딩 저장
                        for idx, chunk_text in enumerate(chunks):
                            # 청크 저장
                            chunk_data = {
                                'document_id': document_id,
                                'content': chunk_text,
                                'chunk_index': idx,
                                'metadata': {
                                    'type': 'internal',
                                    'organization': organization,
                                    'title': doc_data['title'],
                                    'file_path': str(txt_file),
                                    'file_type': 'TXT',
                                    'loaded_at': datetime.now().isoformat()
                                }
                            }
                            chunk_result = self.supabase.client.table('chunks').insert(chunk_data).execute()
                            chunk_id = chunk_result.data[0]['id']
                            
                            # 임베딩 생성 및 저장
                            embedding = self.embeddings.embed_query(chunk_text)
                            embedding_data = {
                                'chunk_id': chunk_id,
                                'embedding': embedding
                            }
                            self.supabase.client.table('embeddings').insert(embedding_data).execute()
                            
                            # Rate limiting
                            time.sleep(0.1)
                        
                        doc_count += 1
                        print(f"    [OK] {txt_file.name} - {len(chunks)} chunks")
                        
                    except Exception as e:
                        print(f"    [ERROR] {txt_file.name}: {str(e)}")
        
        print(f"\n  Total internal documents: {doc_count}")
        return doc_count
    
    def process_external_documents(self):
        """외부 문서 처리"""
        print("\n[3/5] Processing external documents...")
        
        external_dir = Path("./mock_data/external")
        doc_count = 0
        
        for source_dir in external_dir.iterdir():
            if source_dir.is_dir():
                source = source_dir.name.replace('_', ' ')
                print(f"\n  Processing {source}...")
                
                for txt_file in source_dir.glob("*.txt"):
                    try:
                        # 문서 읽기
                        with open(txt_file, 'r', encoding='utf-8') as f:
                            content = f.read()
                        
                        # 카테고리 추출
                        filename_parts = txt_file.stem.split('_')
                        if len(filename_parts) >= 3:
                            category = filename_parts[2]  # 날짜_시간_카테고리_제목
                        else:
                            category = self._extract_category(content)
                        
                        # 문서 정보 생성
                        doc_data = {
                            'type': 'external',
                            'source': source,
                            'title': self._extract_title(content),
                            'category': category,
                            'file_path': str(txt_file),
                            'metadata': {
                                'file_type': 'TXT',
                                'loaded_at': datetime.now().isoformat()
                            }
                        }
                        
                        # 문서 저장
                        doc_result = self.supabase.client.table('documents').insert(doc_data).execute()
                        document_id = doc_result.data[0]['id']
                        
                        # 청킹
                        chunks = self.text_splitter.split_text(content)
                        
                        # 청크 및 임베딩 저장
                        for idx, chunk_text in enumerate(chunks):
                            # 청크 저장
                            chunk_data = {
                                'document_id': document_id,
                                'content': chunk_text,
                                'chunk_index': idx,
                                'metadata': {
                                    'type': 'external',
                                    'source': source,
                                    'category': category,
                                    'title': doc_data['title'],
                                    'file_path': str(txt_file),
                                    'file_type': 'TXT',
                                    'loaded_at': datetime.now().isoformat()
                                }
                            }
                            chunk_result = self.supabase.client.table('chunks').insert(chunk_data).execute()
                            chunk_id = chunk_result.data[0]['id']
                            
                            # 임베딩 생성 및 저장
                            embedding = self.embeddings.embed_query(chunk_text)
                            embedding_data = {
                                'chunk_id': chunk_id,
                                'embedding': embedding
                            }
                            self.supabase.client.table('embeddings').insert(embedding_data).execute()
                            
                            # Rate limiting
                            time.sleep(0.1)
                        
                        doc_count += 1
                        print(f"    [OK] {txt_file.name} - {len(chunks)} chunks")
                        
                    except Exception as e:
                        print(f"    [ERROR] {txt_file.name}: {str(e)}")
        
        print(f"\n  Total external documents: {doc_count}")
        return doc_count
    
    def _extract_category(self, text):
        """텍스트에서 카테고리 추출"""
        categories = ["Macro", "산업", "기술", "리스크", "경쟁사", "정책"]
        
        for category in categories:
            if category in text:
                return category
        
        # 키워드 기반 추출
        if any(word in text for word in ["경제", "금리", "환율", "GDP"]):
            return "Macro"
        elif any(word in text for word in ["배터리", "전기차", "리튬"]):
            return "산업"
        elif any(word in text for word in ["AI", "자동화", "혁신", "전고체"]):
            return "기술"
        elif any(word in text for word in ["규제", "리콜", "화재", "ESG"]):
            return "리스크"
        elif any(word in text for word in ["CATL", "BYD", "Tesla", "LG"]):
            return "경쟁사"
        elif any(word in text for word in ["IRA", "보조금", "정책", "규제"]):
            return "정책"
        
        return "기타"
    
    def _extract_title(self, content):
        """컨텐츠에서 제목 추출"""
        lines = content.split('\n')
        for line in lines:
            if line.strip() and not line.startswith('#'):
                # 첫 번째 의미있는 라인을 제목으로
                return line.strip()[:100]
        return "Untitled"
    
    def verify_upload(self):
        """업로드 검증"""
        print("\n[4/5] Verifying upload...")
        
        try:
            # 문서 수 확인
            docs = self.supabase.client.table('documents').select("type").execute()
            internal_count = len([d for d in docs.data if d['type'] == 'internal'])
            external_count = len([d for d in docs.data if d['type'] == 'external'])
            
            print(f"  - Total documents: {len(docs.data)}")
            print(f"    Internal: {internal_count}")
            print(f"    External: {external_count}")
            
            # 청크 수 확인
            chunks = self.supabase.client.table('chunks').select("id").execute()
            print(f"  - Total chunks: {len(chunks.data)}")
            
            # 임베딩 수 확인
            embeddings = self.supabase.client.table('embeddings').select("id").execute()
            print(f"  - Total embeddings: {len(embeddings.data)}")
            
            if len(docs.data) > 0 and len(chunks.data) > 0 and len(embeddings.data) > 0:
                print("\n  [OK] Upload verified successfully!")
                return True
            else:
                print("\n  [WARNING] Some data may be missing")
                return False
                
        except Exception as e:
            print(f"\n  [ERROR] Verification failed: {str(e)}")
            return False
    
    def run(self):
        """전체 프로세스 실행"""
        print("\n" + "="*70)
        print("MOCK DATA TO SUPABASE UPLOADER")
        print("="*70)
        
        # 1. 기존 데이터 삭제
        if not self.clear_existing_data():
            print("\n[WARNING] Proceeding despite clear errors...")
        
        # 2. 내부 문서 처리
        internal_count = self.process_internal_documents()
        
        # 3. 외부 문서 처리  
        external_count = self.process_external_documents()
        
        # 4. 검증
        self.verify_upload()
        
        # 5. 최종 보고
        print("\n[5/5] Upload Summary")
        print("="*70)
        print(f"Internal documents processed: {internal_count}")
        print(f"External documents processed: {external_count}")
        print(f"Total documents: {internal_count + external_count}")
        print("\n[OK] Mock data upload completed!")
        print("\nYou can now test the RAG system with:")
        print("  py verify_rag_system.py")
        print("  py api_server.py")

if __name__ == "__main__":
    uploader = MockDataUploader()
    uploader.run()