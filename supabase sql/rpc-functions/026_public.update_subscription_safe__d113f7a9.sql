CREATE OR REPLACE FUNCTION public.update_subscription_safe(p_user_id text, p_new_email text, p_new_keyword text)
 RETURNS json
 LANGUAGE plpgsql
AS $function$
declare
  v_conflict_user_id text;
  v_china_time timestamptz := now() + interval '8 hours';
  
  -- 定义变量存储“真正要写入的值”
  v_final_email text;
  v_final_keyword text;
begin
  -- 0. 预处理：先把空字符串变成 NULL，方便后续逻辑统一处理
  p_new_email := nullif(trim(p_new_email), '');
  p_new_keyword := nullif(trim(p_new_keyword), '');

  -- 1. 决定 Email 最终值
  -- 如果传入了新 Email，就用新的；如果没传，就去表里查旧的；如果表里也没（是新用户），那就报错或者接受 NULL
  -- 这里我们简化逻辑：如果 Upsert 时是更新，我们用 COALESCE(新值, EXCLUDED.email) 这种逻辑在下面 SQL 语句里写更方便
  
  -- 但我们需要先拿到“旧值”来做冲突检查
  -- 如果这次没传 Email (p_new_email IS NULL)，那说明用户不想改邮箱，那也就不会有冲突，直接跳过冲突检查
  
  if p_new_email is not null then
    -- 只有当用户真的想改邮箱时，才去查冲突
    select user_id into v_conflict_user_id
    from subscriptions
    where email = p_new_email
    and user_id != p_user_id
    limit 1;

    if v_conflict_user_id is not null then
      return json_build_object(
        'success', false,
        'message', '修改失败：该邮箱已被其他用户绑定。'
      );
    end if;
  end if;

  -- 2. 执行 Upsert (智能合并)
  insert into subscriptions (user_id, email, keyword, status, created_at, updated_at)
  values (
    p_user_id, 
    -- 插入时（新用户）：如果没有传值，就默认为空（或者报错，取决于你的业务，这里假设必须有初始值）
    -- 这里我们假设如果是新用户，必须传入所有值，或者是旧用户更新。
    -- 但 Upsert 语法的精髓在于 ON CONFLICT 下面的 UPDATE
    COALESCE(p_new_email, ''), -- 如果是全新插入但没给email，暂时给个空串避免报错（不太可能发生）
    COALESCE(p_new_keyword, 'AI'), 
    'active', 
    v_china_time, 
    v_china_time
  )
  on conflict (user_id) 
  do update set 
    -- 重点在这里：如果参数不是 NULL，就用参数；如果是 NULL，就保持表里原来的值(subscriptions.email)
    email = COALESCE(p_new_email, subscriptions.email),
    keyword = COALESCE(p_new_keyword, subscriptions.keyword),
    updated_at = v_china_time;

  return json_build_object(
    'success', true,
    'message', '订阅信息已更新！'
  );
end;
$function$;
