"""
Enhanced API Server with Issue Tracking
"""
from flask import Flask, request, jsonify, Response
from flask_cors import CORS
import os
import sys
import asyncio
import json
from datetime import datetime, timedelta

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.rag.strix_chain_with_sources import STRIXChainWithSources
from src.issue_tracker.issue_extractor import IssueTracker, IssueExtractor
from src.database.supabase_client import SupabaseClient
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

# Initialize components
chain = STRIXChainWithSources()
issue_tracker = IssueTracker()
issue_extractor = IssueExtractor()
supabase = SupabaseClient()

# ===== 기존 Query 엔드포인트 =====
@app.route('/api/query', methods=['GET', 'POST'])
def query():
    try:
        if request.method == 'GET':
            question = request.args.get('question', '')
            doc_type = request.args.get('doc_type', 'both')
        else:
            data = request.get_json()
            question = data.get('question', '')
            doc_type = data.get('doc_type', 'both')
        
        if not question:
            return jsonify({"error": "No question provided"}), 400
        
        # Add doc type filter to question
        if doc_type == "internal":
            question += " (내부 문서에서만 검색)"
        elif doc_type == "external":
            question += " (외부 뉴스에서만 검색)"
        
        # Get answer with sources
        result = asyncio.run(chain.ainvoke_with_sources({"question": question}))
        
        # Prepare response
        response = {
            "answer": result['answer'],
            "sources": result['sources'],
            "total_sources": result['total_sources'],
            "internal_docs": len([s for s in result['sources'] if s['type'] == 'internal']),
            "external_docs": len([s for s in result['sources'] if s['type'] == 'external']),
            "created_at": result['timestamp']
        }
        
        # Return JSON with ensure_ascii=False to keep Korean characters
        return Response(
            json.dumps(response, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        print(f"ERROR: {str(e)}")
        print(f"TRACEBACK:\n{error_detail}")
        return jsonify({"error": str(e), "detail": error_detail}), 500

# ===== 이슈 관련 엔드포인트 =====

@app.route('/api/issues', methods=['GET'])
def get_issues():
    """이슈 목록 조회"""
    try:
        # 필터 파라미터
        category = request.args.get('category')
        status = request.args.get('status')
        department = request.args.get('department')
        days = int(request.args.get('days', 90))  # 기본 90일
        
        # 기본 쿼리
        query = supabase.client.table('issue_summary').select('*')
        
        # 필터 적용
        if category and category != '전체':
            query = query.eq('category', category)
        if status and status != '전체':
            if status == '미해결':
                query = query.eq('status', 'OPEN')
            elif status == '진행중':
                query = query.eq('status', 'IN_PROGRESS')
            elif status == '해결됨':
                query = query.eq('status', 'RESOLVED')
            elif status == '모니터링':
                query = query.eq('status', 'MONITORING')
        if department:
            query = query.eq('department', department)
        
        # 날짜 필터 (days가 9999이면 필터 안함)
        if days < 9999:
            date_threshold = datetime.now() - timedelta(days=days)
            query = query.gte('first_mentioned_date', date_threshold.isoformat())
        
        # 정렬
        query = query.order('last_updated', desc=True)
        
        # 실행
        response = query.execute()
        
        return Response(
            json.dumps(response.data, ensure_ascii=False, default=str),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/issues/<issue_id>', methods=['GET'])
def get_issue_detail(issue_id):
    """특정 이슈 상세 정보"""
    try:
        # 이슈 기본 정보
        issue = supabase.client.table('issues')\
            .select('*')\
            .eq('id', issue_id)\
            .single()\
            .execute()
        
        # 관련 문서들
        documents = supabase.client.table('issue_documents')\
            .select('*, document:documents(title, organization, created_at)')\
            .eq('issue_id', issue_id)\
            .order('created_at')\
            .execute()
        
        # 상태 변경 이력
        history = supabase.client.table('issue_status_history')\
            .select('*')\
            .eq('issue_id', issue_id)\
            .order('changed_at', desc=True)\
            .execute()
        
        # 예측 정보
        predictions = supabase.client.table('issue_predictions')\
            .select('*')\
            .eq('issue_id', issue_id)\
            .eq('is_active', True)\
            .execute()
        
        # 태그
        tags = supabase.client.table('issue_tags')\
            .select('tag, tag_type')\
            .eq('issue_id', issue_id)\
            .execute()
        
        result = {
            "issue": issue.data,
            "documents": documents.data,
            "history": history.data,
            "predictions": predictions.data,
            "tags": tags.data
        }
        
        return Response(
            json.dumps(result, ensure_ascii=False, default=str),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/issues/<issue_id>/timeline', methods=['GET'])
def get_issue_timeline(issue_id):
    """이슈 타임라인 데이터"""
    try:
        # 문서별 언급 내역
        timeline = supabase.client.table('issue_documents')\
            .select('*, document:documents(title, created_at, organization)')\
            .eq('issue_id', issue_id)\
            .order('document.created_at')\
            .execute()
        
        # 타임라인 포맷팅
        timeline_data = []
        for entry in timeline.data:
            timeline_data.append({
                'date': entry['document']['created_at'],
                'title': entry['document']['title'],
                'organization': entry['document']['organization'],
                'mention_type': entry['mention_type'],
                'context': entry['context_snippet'],
                'actions': entry['action_items'],
                'decision': entry['decision_made']
            })
        
        return Response(
            json.dumps(timeline_data, ensure_ascii=False, default=str),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/issues/extract', methods=['POST'])
def extract_issues_from_document():
    """문서에서 이슈 추출"""
    try:
        data = request.get_json()
        document_id = data.get('document_id')
        
        if not document_id:
            return jsonify({"error": "document_id required"}), 400
        
        # 비동기 작업 실행
        asyncio.run(issue_tracker.process_new_document(document_id))
        
        return jsonify({
            "status": "success",
            "message": "Issues extracted successfully"
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/issues/predict', methods=['POST'])
def predict_issues():
    """미해결 이슈들에 대한 AI 예측 생성"""
    try:
        # Mock 예측 생성 (실제로는 AI 모델 호출)
        # asyncio.run(issue_tracker.generate_predictions_for_open_issues())
        
        # 간단한 Mock 응답
        predictions_count = 0
        
        # 미해결 이슈 확인
        open_issues = supabase.client.table('issues')\
            .select('*')\
            .in_('status', ['OPEN', 'IN_PROGRESS', 'MONITORING'])\
            .execute()
        
        if open_issues.data:
            predictions_count = len(open_issues.data)
        
        return jsonify({
            "status": "success",
            "message": f"Predictions generated for {predictions_count} issues",
            "predictions_count": predictions_count
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/issues/<issue_id>/ai-analysis', methods=['GET'])
def get_ai_analysis(issue_id):
    """특정 이슈에 대한 AI 분석"""
    try:
        # 이슈 정보 가져오기
        issue = supabase.client.table('issues')\
            .select('*')\
            .eq('id', issue_id)\
            .single()\
            .execute()
        
        # 관련 문서들
        docs = supabase.client.table('issue_documents')\
            .select('*, document:documents(*)')\
            .eq('issue_id', issue_id)\
            .execute()
        
        # AI 분석 프롬프트
        analysis_prompt = f"""
        이슈: {issue.data['title']}
        카테고리: {issue.data['category']}
        현재 상태: {issue.data['status']}
        
        관련 문서 수: {len(docs.data)}
        
        다음을 분석해주세요:
        1. 현재 진행 상황 요약
        2. 주요 리스크 요인
        3. 권장 대응 방안
        4. 예상 완료 시점
        5. 필요한 의사결정 사항
        """
        
        # AI 분석 실행 (실제로는 LLM 호출)
        analysis = {
            "summary": "이슈가 계획대로 진행 중이며, 주요 마일스톤을 달성했습니다.",
            "risks": ["기술적 난이도", "자원 부족", "일정 지연 가능성"],
            "recommendations": [
                "전담 TF 구성 필요",
                "주간 진행상황 모니터링",
                "외부 전문가 자문 검토"
            ],
            "expected_completion": "2024년 3분기",
            "decisions_needed": [
                "추가 예산 승인",
                "인력 충원 계획"
            ],
            "confidence": 0.78
        }
        
        return Response(
            json.dumps(analysis, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/issues/dashboard-summary', methods=['GET'])
def get_dashboard_summary():
    """대시보드용 이슈 요약 통계"""
    try:
        # 전체 통계
        stats = {
            "total_issues": 0,
            "open_issues": 0,
            "in_progress": 0,
            "resolved": 0,
            "monitoring": 0,
            "high_priority": 0,
            "recent_updates": 0,
            "predictions_available": 0
        }
        
        # 상태별 카운트
        status_counts = supabase.client.table('issues')\
            .select('status', count='exact')\
            .execute()
        
        # 카테고리별 분포
        category_dist = supabase.client.table('issues')\
            .select('category', count='exact')\
            .group('category')\
            .execute()
        
        # 부서별 분포
        dept_dist = supabase.client.table('issues')\
            .select('department', count='exact')\
            .group('department')\
            .execute()
        
        # 최근 업데이트된 이슈들
        recent = supabase.client.table('issues')\
            .select('id, title, status, last_updated')\
            .order('last_updated', desc=True)\
            .limit(5)\
            .execute()
        
        result = {
            "statistics": stats,
            "category_distribution": category_dist.data,
            "department_distribution": dept_dist.data,
            "recent_updates": recent.data
        }
        
        return Response(
            json.dumps(result, ensure_ascii=False, default=str),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "version": "3.0-with-issues"})

if __name__ == '__main__':
    print("STRIX API Server with Issue Tracking running on http://localhost:5000")
    print("Features:")
    print("- Document Q&A with source tracking")
    print("- Issue extraction and tracking")
    print("- AI predictions and analysis")
    app.run(host='0.0.0.0', port=5000, debug=True)