CREATE OR REPLACE FUNCTION public.get_news_cache(p_keyword text)
 RETURNS json
 LANGUAGE plpgsql
 SECURITY DEFINER
AS $function$
declare
  v_record record;
  v_ttl_hours int := 1; -- ⏳ 这里设置缓存有效期为 1 小时
begin
  select * into v_record
  from news_cache
  where keyword = lower(trim(p_keyword));

  if v_record is null then
    return json_build_object('hit', false);
  end if;

  -- 检查是否过期
  if (now() - v_record.updated_at) > (v_ttl_hours || ' hours')::interval then
    return json_build_object('hit', false); -- 找到了但过期了，视为未命中
  else
    return json_build_object('hit', true, 'content', v_record.content);
  end if;
end;
$function$;
