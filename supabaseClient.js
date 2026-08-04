import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // eslint-disable-next-line no-console
  console.error(
    "⚠️ لم يتم ضبط VITE_SUPABASE_URL أو VITE_SUPABASE_ANON_KEY.\n" +
    "أنشئ ملف .env (انسخ من .env.example) وضع فيه بيانات مشروعك على Supabase."
  );
}

export const supabase = createClient(
  url || "https://placeholder.supabase.co",
  anonKey || "placeholder-anon-key"
);

export const SUPABASE_READY = Boolean(url && anonKey);
