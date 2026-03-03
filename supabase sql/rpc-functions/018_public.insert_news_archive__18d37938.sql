CREATE OR REPLACE FUNCTION public.insert_news_archive(p_keyword text, p_content text)
 RETURNS json
 LANGUAGE plpgsql
AS $function$
BEGIN
    INSERT INTO daily_news_history (keyword, content)
    VALUES (p_keyword, p_content);
    
    RETURN json_build_object('success', true, 'message', '归档成功');
EXCEPTION WHEN OTHERS THEN
    RETURN json_build_object('success', false, 'message', SQLERRM);
END;
$function$;
