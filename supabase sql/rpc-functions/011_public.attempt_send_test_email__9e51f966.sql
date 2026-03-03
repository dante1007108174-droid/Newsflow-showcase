CREATE OR REPLACE FUNCTION public.attempt_send_test_email(p_user_id text, p_input_email text DEFAULT NULL::text, p_input_keyword text DEFAULT NULL::text)
 RETURNS json
 LANGUAGE plpgsql
 SECURITY DEFINER
AS $function$
declare
  v_final_email text;
  v_final_keyword text;
  v_count int;
  v_limit int := 100; -- 测试期额度
  v_sub_record record;
begin
  -- 1. 查订阅
  select email, keyword into v_sub_record
  from subscriptions
  where user_id = p_user_id
  order by updated_at desc
  limit 1;

  -- 2. 补全邮箱
  if p_input_email is not null and trim(p_input_email) <> '' then
    v_final_email := lower(trim(p_input_email));
  elsif v_sub_record.email is not null then
    v_final_email := v_sub_record.email;
  else
    v_final_email := null;
  end if;

  -- 3. 补全主题
  if p_input_keyword is not null and trim(p_input_keyword) <> '' then
    v_final_keyword := trim(p_input_keyword);
  elsif v_sub_record.keyword is not null then
    v_final_keyword := v_sub_record.keyword;
  else
    v_final_keyword := null;
  end if;

  -- 4. 精准校验反馈
  if v_final_email is null and v_final_keyword is null then
    return json_build_object(
      'allowed', false, 
      'msg', '👋 还需要您提供**邮箱地址**和**主题**（AI、财经、科技），我才能帮您发送测试邮件哦。'
    );
  elsif v_final_email is null then
    return json_build_object(
      'allowed', false, 
      'msg', '👋 请告诉我接收邮件的**邮箱地址**，以便为您发送。'
    );
  elsif v_final_keyword is null then
    return json_build_object(
      'allowed', false, 
      'msg', '👋 请告诉我您想测试哪个**主题**（目前支持：AI、财经、科技）。'
    );
  end if;

  -- 5. 限流检查
  select count(*) into v_count
  from email_logs
  where (user_id = p_user_id or email = v_final_email)
    and created_at > (now() - interval '24 hours');

  if v_count >= v_limit then
    return json_build_object('allowed', false, 'msg', format('🚫 抱歉，今日测试次数已达上限 (%s/%s)。请明天再来试试吧！', v_count, v_limit));
  end if;

  -- 6. 成功
  insert into email_logs (user_id, email, keyword)
  values (p_user_id, v_final_email, v_final_keyword);

  return json_build_object(
    'allowed', true,
    'email', v_final_email,
    'keyword', v_final_keyword,
    'msg', format('✅ 邮件正在发送中，预计 1 分钟左右完成，请留意查收... (今日剩余次数: %s)', v_limit - v_count - 1)
  );
end;
$function$;
