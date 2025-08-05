-- Drop existing function with CASCADE to remove dependencies
DROP FUNCTION IF EXISTS search_chunks CASCADE;

-- Create fixed search_chunks function
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
    doc_type text,
    doc_category text,
    doc_organization text
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
        CAST(d.type AS text),
        CAST(d.category AS text),
        CAST(d.organization AS text)
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