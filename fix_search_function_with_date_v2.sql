-- Drop existing function
DROP FUNCTION IF EXISTS search_chunks CASCADE;

-- Create updated search_chunks function with created_at and correct types
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
    doc_type varchar(50),          -- Changed to varchar(50) to match table
    doc_category varchar(100),      -- Changed to varchar(100) to match table
    doc_organization varchar(100),  -- Changed to varchar(100) to match table
    created_at timestamp           -- Added created_at field
)
LANGUAGE plpgsql
AS $$
BEGIN
    RETURN QUERY
    SELECT 
        c.id as chunk_id,
        c.document_id,
        c.content,
        1 - (e.embedding <=> query_embedding) as similarity,
        c.metadata,
        d.title as doc_title,
        d.type as doc_type,
        d.category as doc_category,
        d.organization as doc_organization,
        d.created_at as created_at
    FROM chunks c
    JOIN embeddings e ON e.chunk_id = c.id
    JOIN documents d ON d.id = c.document_id
    WHERE 
        (1 - (e.embedding <=> query_embedding)) > match_threshold
        AND (filter_type IS NULL OR d.type = filter_type)
        AND (filter_category IS NULL OR d.category = filter_category)
        AND (filter_organization IS NULL OR d.organization = filter_organization)
    ORDER BY similarity DESC
    LIMIT match_count;
END;
$$;

-- Test the function
SELECT 
    doc_title,
    doc_organization,
    created_at,
    doc_type
FROM search_chunks(
    (SELECT embedding FROM embeddings LIMIT 1),
    0.0,
    5
);