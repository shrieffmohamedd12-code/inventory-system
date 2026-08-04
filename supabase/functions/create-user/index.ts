import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  console.log("CREATE USER CALLED");

  if (req.method === "OPTIONS") {
    return new Response("ok", {
      headers: corsHeaders,
    });
  }

  try {
    const { name, pin, role } = await req.json();

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
    );

    const email = `u-${crypto.randomUUID()}@example.com`;

    const { data: userData, error: userError } =
      await admin.auth.admin.createUser({
        email,
        password: pin,
        email_confirm: true,
      });

    if (userError) throw userError;

    const { error: profileError } = await admin
      .from("profiles")
      .insert({
        id: userData.user.id,
        name,
        email,
        role: role || "staff",
      });

    if (profileError) throw profileError;

    return new Response(
      JSON.stringify({
        success: true,
        email,
      }),
      {
        status: 200,
        headers: {
          ...corsHeaders,
          "Content-Type": "application/json",
        },
      }
    );

  } catch (error) {
    console.log("ERROR:", error.message);

    return new Response(
      JSON.stringify({
        error: error.message,
      }),
      {
        status: 500,
        headers: {
          ...corsHeaders,
          "Content-Type": "application/json",
        },
      }
    );
  }
});
