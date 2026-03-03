CREATE OR REPLACE FUNCTION public.force_add_8h_created_at()
 RETURNS trigger
 LANGUAGE plpgsql
AS $function$
BEGIN
  IF TG_OP = 'INSERT' THEN
    IF NEW.created_at IS NULL THEN
      NEW.created_at := now() + INTERVAL '8 hours';
    ELSE
      NEW.created_at := NEW.created_at + INTERVAL '8 hours';
    END IF;
  ELSIF TG_OP = 'UPDATE' THEN
    -- 只有当 created_at 被显式修改才做 +8，避免普通更新重复叠加
    IF NEW.created_at IS DISTINCT FROM OLD.created_at THEN
      NEW.created_at := NEW.created_at + INTERVAL '8 hours';
    END IF;
  END IF;
  RETURN NEW;
END;
$function$;
