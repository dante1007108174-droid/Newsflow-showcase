CREATE OR REPLACE FUNCTION public.set_updated_at()
 RETURNS trigger
 LANGUAGE plpgsql
AS $function$
begin
  new.updated_at = now() at time zone 'Asia/Shanghai';
  return new;
end;
$function$;
