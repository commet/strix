-- STRIX Supabase Setup SQL
-- Run this in Supabase SQL Editor

-- Enable required extensions
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
CREATE EXTENSION IF NOT EXISTS "vector";

-- Documents table
CREATE TABLE IF NOT EXISTS documents (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    type VARCHAR(50) NOT NULL CHECK (type IN ('internal', 'external')),
    source VARCHAR(255),
    title TEXT,
    organization VARCHAR(100),
    category VARCHAR(100),
    created_at TIMESTAMP DEFAULT NOW(),
    file_path TEXT,
    metadata JSONB,
    CONSTRAINT unique_file_path UNIQUE (file_path)
);

-- Create indexes
CREATE INDEX idx_documents_type ON documents(type);
CREATE INDEX idx_documents_organization ON documents(organization);
CREATE INDEX idx_documents_category ON documents(category);
CREATE INDEX idx_documents_created_at ON documents(created_at);

-- Chunks table
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

CREATE INDEX idx_chunks_document_id ON chunks(document_id);

-- Embeddings table
CREATE TABLE IF NOT EXISTS embeddings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    chunk_id UUID REFERENCES chunks(id) ON DELETE CASCADE,
    embedding vector(1536),
    model VARCHAR(50) DEFAULT 'text-embedding-ada-002',
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX idx_embeddings_chunk_id ON embeddings(chunk_id);

-- Correlations table
CREATE TABLE IF NOT EXISTS correlations (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    internal_doc_id UUID REFERENCES documents(id),
    external_doc_id UUID REFERENCES documents(id),
    score FLOAT CHECK (score >= 0 AND score <= 1),
    reasoning TEXT,
    verified BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX idx_correlations_internal ON correlations(internal_doc_id);
CREATE INDEX idx_correlations_external ON correlations(external_doc_id);
CREATE INDEX idx_correlations_score ON correlations(score DESC);

-- Search logs table
CREATE TABLE IF NOT EXISTS search_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    query TEXT,
    results JSONB,
    user_feedback JSONB,
    created_at TIMESTAMP DEFAULT NOW()
);

-- Keyword learning table
CREATE TABLE IF NOT EXISTS keyword_learning (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    category VARCHAR(50),
    keyword VARCHAR(100),
    frequency INTEGER DEFAULT 1,
    relevance_score FLOAT,
    status VARCHAR(20) DEFAULT 'pending' CHECK (status IN ('pending', 'approved', 'rejected')),
    created_at TIMESTAMP DEFAULT NOW(),
    CONSTRAINT unique_category_keyword UNIQUE (category, keyword)
);

CREATE INDEX idx_keyword_learning_category ON keyword_learning(category);
CREATE INDEX idx_keyword_learning_status ON keyword_learning(status);

-- Function for vector similarity search
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
        c.id as chunk_id,
        c.document_id,
        c.content,
        1 - (e.embedding <=> query_embedding) as similarity,
        c.metadata,
        d.title as doc_title,
        d.type as doc_type,
        d.category as doc_category,
        d.organization as doc_organization
    FROM chunks c
    JOIN embeddings e ON e.chunk_id = c.id
    JOIN documents d ON d.id = c.document_id
    WHERE 
        1 - (e.embedding <=> query_embedding) > match_threshold
        AND (filter_type IS NULL OR d.type = filter_type)
        AND (filter_category IS NULL OR d.category = filter_category)
        AND (filter_organization IS NULL OR d.organization = filter_organization)
    ORDER BY similarity DESC
    LIMIT match_count;
END;
$$;

-- Function to get document statistics
CREATE OR REPLACE FUNCTION get_document_stats()
RETURNS TABLE (
    stat_type text,
    stat_value bigint
)
LANGUAGE plpgsql
AS $$
BEGIN
    RETURN QUERY
    SELECT 'total_documents'::text, COUNT(*)::bigint FROM documents
    UNION ALL
    SELECT 'internal_documents'::text, COUNT(*)::bigint FROM documents WHERE type = 'internal'
    UNION ALL
    SELECT 'external_documents'::text, COUNT(*)::bigint FROM documents WHERE type = 'external'
    UNION ALL
    SELECT 'total_chunks'::text, COUNT(*)::bigint FROM chunks
    UNION ALL
    SELECT 'total_correlations'::text, COUNT(*)::bigint FROM correlations;
END;
$$;

-- Grant permissions (adjust based on your needs)
GRANT ALL ON ALL TABLES IN SCHEMA public TO authenticated;
GRANT ALL ON ALL FUNCTIONS IN SCHEMA public TO authenticated;
GRANT ALL ON ALL SEQUENCES IN SCHEMA public TO authenticated;