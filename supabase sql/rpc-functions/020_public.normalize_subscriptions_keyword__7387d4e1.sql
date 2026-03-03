CREATE OR REPLACE FUNCTION public.normalize_subscriptions_keyword()
 RETURNS trigger
 LANGUAGE plpgsql
AS $function$
begin
  if new.keyword is null then
    return new;
  end if;

  -- trim
  new.keyword := btrim(new.keyword);

  -- normalize AI variants
  if lower(new.keyword) = 'ai' or new.keyword = '人工智能' then
    new.keyword := 'ai';
  end if;

  return new;
end;
$function$;
