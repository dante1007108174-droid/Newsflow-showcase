CREATE OR REPLACE FUNCTION public.get_latest_daily_news_html(p_keyword text, p_max_age_hours integer DEFAULT 24)
 RETURNS TABLE(found boolean, html_content text, created_at timestamp with time zone)
 LANGUAGE plpgsql
 SECURITY DEFINER
AS $function$
DECLARE
  v_search_key text;
BEGIN
  v_search_key := lower(trim(COALESCE(p_keyword, '')));

  IF v_search_key = '人工智能' THEN
    v_search_key := 'ai';
  END IF;

  -- 只在找到时返回，没找到就不返回任何行（body 就是 []）
  RETURN QUERY
  SELECT
    true,
    d.content,
    d.created_at
  FROM daily_news_history d
  WHERE d.keyword = v_search_key
    AND d.created_at > (now() - (p_max_age_hours || ' hours')::interval)
  ORDER BY d.created_at DESC
  LIMIT 1;

  -- 删掉了 IF NOT FOUND 那段，这样没缓存时返回空数组 []
END;
$function$;
