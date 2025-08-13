-- Issue Tracking System Tables for STRIX

-- 1. 이슈 마스터 테이블
CREATE TABLE IF NOT EXISTS issues (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    issue_key VARCHAR(50) UNIQUE, -- 예: ISS-2024-001
    title TEXT NOT NULL,
    category VARCHAR(100), -- 전략, 기술, 리스크, 경쟁사, 정책 등
    priority VARCHAR(20), -- HIGH, MEDIUM, LOW
    status VARCHAR(50), -- OPEN, IN_PROGRESS, RESOLVED, MONITORING
    first_mentioned_date DATE,
    last_updated DATE,
    resolution_date DATE,
    department VARCHAR(100), -- 주관 부서
    owner VARCHAR(100), -- 담당자
    description TEXT,
    impact_assessment TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    metadata JSONB
);

-- 2. 이슈-문서 연결 테이블 (이슈가 언급된 문서들)
CREATE TABLE IF NOT EXISTS issue_documents (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    issue_id UUID REFERENCES issues(id) ON DELETE CASCADE,
    document_id UUID REFERENCES documents(id) ON DELETE CASCADE,
    mention_type VARCHAR(50), -- FIRST_MENTION, UPDATE, RESOLUTION, REFERENCE
    context_snippet TEXT, -- 해당 문서에서 이슈가 언급된 부분
    action_items TEXT[], -- 해당 문서에서 제시된 액션 아이템들
    decision_made TEXT, -- 해당 문서에서 내려진 결정사항
    created_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(issue_id, document_id)
);

-- 3. 이슈 상태 변경 이력
CREATE TABLE IF NOT EXISTS issue_status_history (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    issue_id UUID REFERENCES issues(id) ON DELETE CASCADE,
    old_status VARCHAR(50),
    new_status VARCHAR(50),
    changed_by VARCHAR(100),
    change_reason TEXT,
    document_id UUID REFERENCES documents(id), -- 상태 변경의 근거가 된 문서
    changed_at TIMESTAMP DEFAULT NOW()
);

-- 4. 이슈 간 연관관계
CREATE TABLE IF NOT EXISTS issue_relationships (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    parent_issue_id UUID REFERENCES issues(id) ON DELETE CASCADE,
    related_issue_id UUID REFERENCES issues(id) ON DELETE CASCADE,
    relationship_type VARCHAR(50), -- DEPENDS_ON, BLOCKS, RELATED, DUPLICATE
    created_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(parent_issue_id, related_issue_id)
);

-- 5. 이슈 예측 및 대응방안
CREATE TABLE IF NOT EXISTS issue_predictions (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    issue_id UUID REFERENCES issues(id) ON DELETE CASCADE,
    prediction_type VARCHAR(50), -- RISK, OPPORTUNITY, DEADLINE
    prediction_content TEXT,
    confidence_score FLOAT, -- 0-1
    recommended_actions TEXT[],
    ai_reasoning TEXT,
    predicted_at TIMESTAMP DEFAULT NOW(),
    is_active BOOLEAN DEFAULT TRUE
);

-- 6. 이슈 태그 (키워드)
CREATE TABLE IF NOT EXISTS issue_tags (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    issue_id UUID REFERENCES issues(id) ON DELETE CASCADE,
    tag VARCHAR(100),
    tag_type VARCHAR(50), -- KEYWORD, TECHNOLOGY, COMPETITOR, REGULATION
    created_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(issue_id, tag)
);

-- 인덱스 생성
CREATE INDEX idx_issues_status ON issues(status);
CREATE INDEX idx_issues_category ON issues(category);
CREATE INDEX idx_issues_department ON issues(department);
CREATE INDEX idx_issue_documents_issue_id ON issue_documents(issue_id);
CREATE INDEX idx_issue_documents_document_id ON issue_documents(document_id);
CREATE INDEX idx_issue_tags_tag ON issue_tags(tag);

-- 뷰: 이슈별 최신 상태 요약
CREATE OR REPLACE VIEW issue_summary AS
SELECT 
    i.id,
    i.issue_key,
    i.title,
    i.category,
    i.priority,
    i.status,
    i.department,
    i.owner,
    i.first_mentioned_date,
    i.last_updated,
    COUNT(DISTINCT id.document_id) as document_count,
    COUNT(DISTINCT it.tag) as tag_count,
    ARRAY_AGG(DISTINCT it.tag) FILTER (WHERE it.tag IS NOT NULL) as tags,
    MAX(ip.prediction_content) as latest_prediction
FROM issues i
LEFT JOIN issue_documents id ON i.id = id.issue_id
LEFT JOIN issue_tags it ON i.id = it.issue_id
LEFT JOIN issue_predictions ip ON i.id = ip.issue_id AND ip.is_active = TRUE
GROUP BY i.id, i.issue_key, i.title, i.category, i.priority, i.status, 
         i.department, i.owner, i.first_mentioned_date, i.last_updated;