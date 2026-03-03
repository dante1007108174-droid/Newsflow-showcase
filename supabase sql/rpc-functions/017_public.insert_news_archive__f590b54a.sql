CREATE OR REPLACE FUNCTION public.insert_news_archive(json_payload text)
 RETURNS json
 LANGUAGE plpgsql
AS $function$
DECLARE
    payload JSON;
BEGIN
    payload := json_payload::JSON;
    
    INSERT INTO daily_news_history (keyword, content)
    VALUES (payload->>'p_keyword', payload->>'p_content');
    
    RETURN json_build_object('success', true, 'message', '归档成功');
EXCEPTION WHEN OTHERS THEN
    RETURN json_build_object('success', false, 'message', SQLERRM);
END;
$function$;
