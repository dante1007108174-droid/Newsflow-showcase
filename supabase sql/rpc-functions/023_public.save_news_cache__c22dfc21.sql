CREATE OR REPLACE FUNCTION public.save_news_cache(p_keyword text, p_content text)
 RETURNS void
 LANGUAGE plpgsql
 SECURITY DEFINER
AS $function$
begin
  insert into news_cache (keyword, content, updated_at)
  values (lower(trim(p_keyword)), p_content, now())
  on conflict (keyword) 
  do update set 
    content = excluded.content,
    updated_at = excluded.updated_at;
end;
$function$;
