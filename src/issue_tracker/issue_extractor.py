"""
Issue Extraction and Analysis Engine
"""
import re
from typing import List, Dict, Any, Tuple
from datetime import datetime
import asyncio
from langchain_openai import ChatOpenAI
from langchain.prompts import ChatPromptTemplate
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from src.database.supabase_client import SupabaseClient


class IssueExtractor:
    def __init__(self):
        self.llm = ChatOpenAI(temperature=0.1, model="gpt-4o-mini")
        self.supabase = SupabaseClient()
        
        # 이슈 추출 프롬프트
        self.extraction_prompt = ChatPromptTemplate.from_messages([
            ("system", """당신은 기업 문서에서 주요 이슈를 추출하는 전문 분석가입니다.
            
            다음 기준으로 이슈를 추출하세요:
            1. 해결이 필요한 문제나 과제
            2. 의사결정이 필요한 사항
            3. 리스크나 기회 요인
            4. 전략적 검토가 필요한 사항
            5. Follow-up이 필요한 액션 아이템
            
            각 이슈에 대해 다음을 파악하세요:
            - 이슈 제목
            - 카테고리 (전략/기술/리스크/경쟁사/정책/운영)
            - 우선순위 (HIGH/MEDIUM/LOW)
            - 현재 상태 (언급만 됨/검토 중/대응 중/해결됨)
            - 관련 부서
            - 핵심 키워드
            - 액션 아이템 (있다면)
            - 의사결정 사항 (있다면)
            """),
            ("user", """문서 제목: {title}
            문서 날짜: {date}
            문서 내용:
            {content}
            
            위 문서에서 이슈를 추출하고 JSON 형식으로 반환하세요:
            {{
                "issues": [
                    {{
                        "title": "이슈 제목",
                        "category": "카테고리",
                        "priority": "우선순위",
                        "status": "상태",
                        "department": "관련 부서",
                        "keywords": ["키워드1", "키워드2"],
                        "context_snippet": "이슈가 언급된 문맥",
                        "action_items": ["액션1", "액션2"],
                        "decision": "결정사항"
                    }}
                ]
            }}""")
        ])
        
        # 이슈 연관성 분석 프롬프트
        self.relationship_prompt = ChatPromptTemplate.from_messages([
            ("system", "두 이슈 간의 연관성을 분석하는 전문가입니다."),
            ("user", """이슈 1: {issue1_title}
            설명: {issue1_desc}
            
            이슈 2: {issue2_title}
            설명: {issue2_desc}
            
            두 이슈의 관계를 분석하세요:
            - DEPENDS_ON: 이슈1이 이슈2에 의존
            - BLOCKS: 이슈1이 이슈2를 막음
            - RELATED: 관련 있음
            - DUPLICATE: 중복
            - NONE: 관계 없음
            
            관계 타입과 이유를 JSON으로 반환하세요.""")
        ])
    
    async def extract_issues_from_document(self, document: Dict[str, Any]) -> List[Dict[str, Any]]:
        """문서에서 이슈 추출"""
        try:
            # LLM을 사용하여 이슈 추출
            response = await self.llm.ainvoke(
                self.extraction_prompt.format_messages(
                    title=document.get('title', ''),
                    date=document.get('created_at', ''),
                    content=document.get('content', '')[:3000]  # 토큰 제한
                )
            )
            
            # JSON 파싱
            import json
            result = json.loads(response.content)
            
            extracted_issues = []
            for issue_data in result.get('issues', []):
                # 이슈 키 생성
                issue_key = self._generate_issue_key(issue_data['title'])
                
                issue = {
                    'issue_key': issue_key,
                    'title': issue_data['title'],
                    'category': issue_data['category'],
                    'priority': issue_data['priority'],
                    'status': self._map_status(issue_data['status']),
                    'department': issue_data['department'],
                    'description': issue_data.get('context_snippet', ''),
                    'first_mentioned_date': document.get('created_at'),
                    'last_updated': document.get('created_at'),
                    'metadata': {
                        'keywords': issue_data.get('keywords', []),
                        'extracted_from': document.get('id')
                    }
                }
                
                extracted_issues.append({
                    'issue': issue,
                    'document_link': {
                        'document_id': document.get('id'),
                        'mention_type': 'FIRST_MENTION',
                        'context_snippet': issue_data.get('context_snippet', ''),
                        'action_items': issue_data.get('action_items', []),
                        'decision_made': issue_data.get('decision', '')
                    },
                    'tags': issue_data.get('keywords', [])
                })
            
            return extracted_issues
            
        except Exception as e:
            print(f"Error extracting issues: {str(e)}")
            return []
    
    def _generate_issue_key(self, title: str) -> str:
        """이슈 키 생성"""
        # 제목에서 영문자만 추출
        clean_title = re.sub(r'[^a-zA-Z가-힣]', '', title)
        prefix = clean_title[:3].upper() if clean_title else 'ISS'
        
        # 타임스탬프 기반 고유 번호
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        return f"{prefix}-{timestamp}"
    
    def _map_status(self, status_text: str) -> str:
        """상태 텍스트를 표준 상태로 매핑"""
        status_map = {
            '언급만 됨': 'OPEN',
            '검토 중': 'IN_PROGRESS',
            '대응 중': 'IN_PROGRESS',
            '해결됨': 'RESOLVED',
            '모니터링': 'MONITORING'
        }
        return status_map.get(status_text, 'OPEN')
    
    async def analyze_issue_relationships(self, issue1: Dict, issue2: Dict) -> Dict:
        """두 이슈 간의 관계 분석"""
        response = await self.llm.ainvoke(
            self.relationship_prompt.format_messages(
                issue1_title=issue1['title'],
                issue1_desc=issue1.get('description', ''),
                issue2_title=issue2['title'],
                issue2_desc=issue2.get('description', '')
            )
        )
        
        import json
        return json.loads(response.content)
    
    async def predict_issue_development(self, issue: Dict, related_docs: List[Dict]) -> Dict:
        """이슈의 향후 전개 예측"""
        prediction_prompt = ChatPromptTemplate.from_messages([
            ("system", """당신은 기업 이슈 분석 전문가입니다.
            과거 문서들을 바탕으로 이슈의 향후 전개를 예측하고 대응방안을 제시하세요."""),
            ("user", """이슈: {issue_title}
            현재 상태: {status}
            카테고리: {category}
            
            관련 문서 이력:
            {doc_history}
            
            다음을 예측하고 제안하세요:
            1. 향후 전개 시나리오
            2. 리스크 요인
            3. 기회 요인
            4. 권장 액션 아이템
            5. 예상 해결 시기
            
            JSON 형식으로 반환하세요.""")
        ])
        
        # 문서 이력 포맷팅
        doc_history = "\n".join([
            f"- {doc['created_at']}: {doc['title']}" 
            for doc in related_docs
        ])
        
        response = await self.llm.ainvoke(
            prediction_prompt.format_messages(
                issue_title=issue['title'],
                status=issue['status'],
                category=issue['category'],
                doc_history=doc_history
            )
        )
        
        import json
        prediction = json.loads(response.content)
        
        return {
            'prediction_type': 'COMPREHENSIVE',
            'prediction_content': prediction.get('scenario', ''),
            'confidence_score': 0.75,  # 실제로는 더 정교한 계산 필요
            'recommended_actions': prediction.get('actions', []),
            'ai_reasoning': str(prediction)
        }


class IssueTracker:
    """이슈 추적 및 관리 시스템"""
    
    def __init__(self):
        self.supabase = SupabaseClient()
        self.extractor = IssueExtractor()
    
    async def process_new_document(self, document_id: str):
        """새 문서 처리 및 이슈 추출"""
        # 문서 가져오기
        doc = self.supabase.get_document_by_id(document_id)
        if not doc:
            return
        
        # 문서 내용 가져오기 (chunks에서)
        chunks = self.supabase.client.table('chunks')\
            .select('content')\
            .eq('document_id', document_id)\
            .order('chunk_index')\
            .execute()
        
        doc['content'] = ' '.join([chunk['content'] for chunk in chunks.data])
        
        # 이슈 추출
        extracted_issues = await self.extractor.extract_issues_from_document(doc)
        
        for issue_data in extracted_issues:
            # 기존 이슈 확인
            existing = self.supabase.client.table('issues')\
                .select('id')\
                .eq('title', issue_data['issue']['title'])\
                .execute()
            
            if existing.data:
                # 기존 이슈 업데이트
                issue_id = existing.data[0]['id']
                self._update_existing_issue(issue_id, issue_data, doc)
            else:
                # 새 이슈 생성
                self._create_new_issue(issue_data)
    
    def _create_new_issue(self, issue_data: Dict):
        """새 이슈 생성"""
        # 이슈 생성
        issue_response = self.supabase.client.table('issues')\
            .insert(issue_data['issue'])\
            .execute()
        
        issue_id = issue_response.data[0]['id']
        
        # 문서 연결
        doc_link = issue_data['document_link']
        doc_link['issue_id'] = issue_id
        self.supabase.client.table('issue_documents')\
            .insert(doc_link)\
            .execute()
        
        # 태그 추가
        for tag in issue_data['tags']:
            self.supabase.client.table('issue_tags')\
                .insert({
                    'issue_id': issue_id,
                    'tag': tag,
                    'tag_type': 'KEYWORD'
                })\
                .execute()
    
    def _update_existing_issue(self, issue_id: str, issue_data: Dict, document: Dict):
        """기존 이슈 업데이트"""
        # 마지막 업데이트 날짜 갱신
        self.supabase.client.table('issues')\
            .update({'last_updated': document['created_at']})\
            .eq('id', issue_id)\
            .execute()
        
        # 문서 연결 추가
        doc_link = issue_data['document_link']
        doc_link['issue_id'] = issue_id
        doc_link['mention_type'] = 'UPDATE'
        
        self.supabase.client.table('issue_documents')\
            .upsert(doc_link)\
            .execute()
    
    async def generate_predictions_for_open_issues(self):
        """미해결 이슈들에 대한 예측 생성"""
        # 미해결 이슈 조회
        open_issues = self.supabase.client.table('issues')\
            .select('*')\
            .in_('status', ['OPEN', 'IN_PROGRESS', 'MONITORING'])\
            .execute()
        
        for issue in open_issues.data:
            # 관련 문서들 조회
            related_docs = self.supabase.client.table('issue_documents')\
                .select('*, document:documents(*)')\
                .eq('issue_id', issue['id'])\
                .order('created_at')\
                .execute()
            
            docs = [rd['document'] for rd in related_docs.data]
            
            # 예측 생성
            prediction = await self.extractor.predict_issue_development(issue, docs)
            prediction['issue_id'] = issue['id']
            
            # 기존 예측 비활성화
            self.supabase.client.table('issue_predictions')\
                .update({'is_active': False})\
                .eq('issue_id', issue['id'])\
                .execute()
            
            # 새 예측 저장
            self.supabase.client.table('issue_predictions')\
                .insert(prediction)\
                .execute()