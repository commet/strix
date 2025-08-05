-- 1. 현재 search_chunks 함수 확인
SELECT 
    proname as function_name,
    proargtypes,
    prorettype,
    prosrc
FROM pg_proc 
WHERE proname = 'search_chunks';

-- 2. 함수의 반환 타입 확인
SELECT 
    a.attname as column_name,
    t.typname as data_type,
    a.attnum as position
FROM pg_attribute a
JOIN pg_type t ON a.atttypid = t.oid
WHERE a.attrelid = (
    SELECT prorettype 
    FROM pg_proc 
    WHERE proname = 'search_chunks'
)
ORDER BY a.attnum;