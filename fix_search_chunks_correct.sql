-- 기존 함수 완전히 삭제
DROP FUNCTION IF EXISTS search_chunks(vector, float, int, text, text, text);

-- 정확한 타입으로 함수 재생성
CREATE OR REPLACE FUNCTION search_chunks(
    query_embedding vector(1536),
    match_threshold float DEFAULT 0.7,
    match_count int DEFAULT 10,
    filter_type text DEFAULT NULL,
    filter_category text DEFAULT NULL,
    filter_organization text DEFAULT NULL
)
RETURNS TABLE (
    chunk_id uuid,
    document_id uuid,
    content text,
    similarity float,
    metadata jsonb,
    doc_title text,
    doc_type VARCHAR(50),     -- 정확한 타입 지정
    doc_category VARCHAR(50),  -- 정확한 타입 지정
    doc_organization VARCHAR(50) -- 정확한 타입 지정
)
LANGUAGE plpgsql
AS $$
BEGIN
    RETURN QUERY
    SELECT 
        c.id,
        c.document_id,
        c.content,
        1 - (e.embedding <=> query_embedding),
        c.metadata,
        d.title,
        d.type,        -- 캐스팅 제거
        d.category,    -- 캐스팅 제거
        d.organization -- 캐스팅 제거
    FROM chunks c
    JOIN embeddings e ON c.id = e.chunk_id
    JOIN documents d ON c.document_id = d.id
    WHERE 
        1 - (e.embedding <=> query_embedding) > match_threshold
        AND (filter_type IS NULL OR d.type = filter_type)
        AND (filter_category IS NULL OR d.category = filter_category)
        AND (filter_organization IS NULL OR d.organization = filter_organization)
    ORDER BY 4 DESC
    LIMIT match_count;
END;
$$;