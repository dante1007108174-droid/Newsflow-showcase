CREATE OR REPLACE FUNCTION public.get_user_email(p_user_id text)
 RETURNS TABLE(email text)
 LANGUAGE plpgsql
 SECURITY DEFINER
AS $function$
BEGIN
  RETURN QUERY
  SELECT s.email
  FROM subscriptions s
  WHERE s.user_id = p_user_id
  LIMIT 1;
END;
$function$;
