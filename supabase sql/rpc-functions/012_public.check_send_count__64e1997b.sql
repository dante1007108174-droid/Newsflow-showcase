CREATE OR REPLACE FUNCTION public.check_send_count(p_user_id text, p_limit integer DEFAULT 99)
 RETURNS TABLE(allowed boolean, current_count integer)
 LANGUAGE plpgsql
 SECURITY DEFINER
AS $function$
DECLARE
  v_count int;
BEGIN
  SELECT COUNT(*)::int INTO v_count
  FROM daily_email_logs
  WHERE user_id = p_user_id
    AND created_at > now() - interval '24 hours';

  RETURN QUERY SELECT v_count < p_limit, v_count;
END;
$function$;
