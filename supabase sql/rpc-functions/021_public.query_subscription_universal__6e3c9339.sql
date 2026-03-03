CREATE OR REPLACE FUNCTION public.query_subscription_universal(p_user_id text, p_target_email text)
 RETURNS json
 LANGUAGE plpgsql
AS $function$
declare
  -- 关键修改：明确指定类型，或者手动初始化
  v_record subscriptions%ROWTYPE; 
begin
  -- 0. 参数清洗
  p_user_id := nullif(trim(p_user_id), '');
  p_target_email := nullif(trim(p_target_email), '');

  -- 1. 🔍 定位 (Find)
  
  -- 尝试按 email 找
  if p_target_email is not null then
    select * into v_record from subscriptions where email = p_target_email limit 1;
  end if;

  -- 如果上面没找到（v_record.id 是 null），再尝试按 user_id 找
  if v_record.id is null and p_user_id is not null then
    select * into v_record from subscriptions where user_id = p_user_id limit 1;
  end if;

  -- 2. 📦 返回结果
  -- 只有当 ID 真的有值时，才算找到了
  if v_record.id is not null then
    return json_build_object(
      'success', true,
      'data', json_build_object(
        'email', v_record.email,
        'keyword', v_record.keyword,
        'status', v_record.status,
        'updated_at', to_char(v_record.updated_at, 'YYYY-MM-DD HH24:MI:SS')
      )
    );
  else
    return json_build_object(
      'success', false,
      'message', '未找到相关的订阅记录。'
    );
  end if;
end;
$function$;
