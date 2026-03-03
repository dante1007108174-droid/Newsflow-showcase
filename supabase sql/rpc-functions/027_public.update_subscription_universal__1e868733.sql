CREATE OR REPLACE FUNCTION public.update_subscription_universal(p_user_id text, p_target_email text, p_new_email text, p_new_keyword text)
 RETURNS json
 LANGUAGE plpgsql
AS $function$
declare
  v_record_id uuid;
  v_conflict_check text;
  v_china_time timestamptz := now() + interval '8 hours';
begin
  -- 0. 参数清洗
  p_user_id := nullif(trim(p_user_id), '');
  p_target_email := nullif(trim(p_target_email), '');
  p_new_email := nullif(trim(p_new_email), '');
  p_new_keyword := nullif(trim(p_new_keyword), '');

  -- ========== ✅ 修改：主题归一化 + 白名单校验 ==========
  if p_new_keyword is not null then
    -- 统一清洗
    p_new_keyword := btrim(p_new_keyword);

    -- 归一化：AI/ai/Ai/人工智能 -> ai（小写）
    if lower(p_new_keyword) = 'ai' or p_new_keyword = '人工智能' then
      p_new_keyword := 'ai';
    end if;

    -- 白名单校验：只允许 ai、科技、财经
    if p_new_keyword not in ('ai', '科技', '财经') then
      return json_build_object(
        'success', false,
        'message', format('不支持的订阅主题「%s」，仅支持：ai、科技、财经。', p_new_keyword)
      );
    end if;
  end if;
  -- ========== ✅ 结束 ==========

  -- 1. 🔍 定位 (Find)
  if p_target_email is not null then
    select id into v_record_id from subscriptions where email = p_target_email limit 1;
  end if;
  if v_record_id is null and p_user_id is not null then
    select id into v_record_id from subscriptions where user_id = p_user_id limit 1;
  end if;

  -- 2. 🛡️ 冲突检查 (Conflict)
  if p_new_email is not null then
    select user_id into v_conflict_check
    from subscriptions
    where email = p_new_email
    and id != coalesce(v_record_id, '00000000-0000-0000-0000-000000000000'::uuid)
    limit 1;

    if v_conflict_check is not null then
      return json_build_object('success', false, 'message', '修改失败：新邮箱已被其他账号占用。');
    end if;
  end if;

  -- 3. 📝 执行 (Action)
  if v_record_id is not null then
    -- A. 老用户更新 (允许部分更新，不查必填)
    -- ========== 🆕 修改：加上 status = 'active' ==========
    update subscriptions set
      email = coalesce(p_new_email, email),
      keyword = coalesce(p_new_keyword, keyword),
      status = 'active',  -- 🆕 恢复为 active
      updated_at = v_china_time
    where id = v_record_id;
    return json_build_object('success', true, 'message', '订阅信息更新成功！');
    -- ========== 🆕 结束 ==========

  else
    -- B. 新用户创建 (必须全都有！)
    -- 🛑 校验开始
    if p_new_email is null then
        return json_build_object('success', false, 'message', '首次订阅失败：请提供邮箱地址。');
    end if;

    if p_new_keyword is null then
        return json_build_object('success', false, 'message', '首次订阅失败：请提供订阅主题关键词。');
    end if;
    -- 🛑 校验结束

    insert into subscriptions (user_id, email, keyword, status, created_at, updated_at)
    values (p_user_id, p_new_email, p_new_keyword, 'active', v_china_time, v_china_time);

    return json_build_object('success', true, 'message', '新订阅创建成功！');
  end if;
end;
$function$;
