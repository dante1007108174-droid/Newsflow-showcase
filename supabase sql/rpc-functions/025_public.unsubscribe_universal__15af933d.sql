CREATE OR REPLACE FUNCTION public.unsubscribe_universal(p_user_id text, p_target_email text)
 RETURNS json
 LANGUAGE plpgsql
AS $function$
declare
  v_record_id uuid;
  v_email_found text;
begin
  -- 0. 参数清洗
  p_user_id := nullif(trim(p_user_id), '');
  p_target_email := nullif(trim(p_target_email), '');

  -- 1. 🔍 定位目标 (逻辑复用)
  if p_target_email is not null then
    select id, email into v_record_id, v_email_found from subscriptions where email = p_target_email limit 1;
  end if;

  if v_record_id is null and p_user_id is not null then
    select id, email into v_record_id, v_email_found from subscriptions where user_id = p_user_id limit 1;
  end if;

  -- 2. 执行取消
  if v_record_id is not null then
    -- 软删除：设置 status = 'inactive'
    -- 如果你想硬删除，改成: delete from subscriptions where id = v_record_id;
    update subscriptions 
    set status = 'inactive', updated_at = (now() + interval '8 hours')
    where id = v_record_id;
    
    return json_build_object(
      'success', true, 
      'message', format('已成功取消 %s 的订阅推送。', coalesce(v_email_found, ''))
    );
  else
    return json_build_object(
      'success', false, 
      'message', '未找到有效的活跃订阅，无需取消。'
    );
  end if;
end;
$function$;
