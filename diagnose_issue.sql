-- 1. documents 테이블 구조 확인
SELECT 
    column_name,
    data_type,
    character_maximum_length
FROM information_schema.columns
WHERE table_name = 'documents'
ORDER BY ordinal_position;

-- 2. 현재 search_chunks 함수의 정확한 정의 확인
\df+ search_chunks