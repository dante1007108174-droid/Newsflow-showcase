CREATE OR REPLACE FUNCTION public.log_email_send(p_user_id text, p_email text, p_keyword text)
 RETURNS void
 LANGUAGE plpgsql
 SECURITY DEFINER
AS $function$
BEGIN
  INSERT INTO daily_email_logs (user_id, email, keyword)
  VALUES (p_user_id, p_email, p_keyword);
END;
$function$;
