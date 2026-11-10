package uy.cg.avisos;

import android.content.Context;
import android.content.SharedPreferences;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.regex.Pattern;

/** Manda cada aviso de movimiento al receptor de la app de gastos (Google Apps Script). */
public final class Sender {
    private Sender() {}

    static final Pattern MONTO = Pattern.compile("(U\\$S|US\\$|USD|UYU|\\$)\\s?-?\\d", Pattern.CASE_INSENSITIVE);
    static final Pattern CLAVE = Pattern.compile(
            "compra|consumo|d[eé]bito|transacci|transfer|acredit|pago|retiro|cobro|dep[oó]sito|cargo|operaci",
            Pattern.CASE_INSENSITIVE);

    static boolean pareceMovimiento(String t) {
        return t != null && MONTO.matcher(t).find() && CLAVE.matcher(t).find();
    }

    static SharedPreferences prefs(Context c) {
        return c.getSharedPreferences("cfg", Context.MODE_PRIVATE);
    }

    /** canal: APP o SMS. */
    static void enviar(Context ctx, String canal, String app, String titulo, String texto, Runnable alTerminar) {
        final Context c = ctx.getApplicationContext();
        final SharedPreferences p = prefs(c);
        final String url = p.getString("url", ""), k = p.getString("k", "");
        if (url.isEmpty() || k.isEmpty()) {
            log(c, canal + " " + app + ": falta configurar");
            if (alTerminar != null) alTerminar.run();
            return;
        }
        final String body;
        try {
            JSONObject o = new JSONObject();
            o.put("k", k);
            o.put("app", canal + "|" + app);
            o.put("titulo", titulo == null ? "" : titulo);
            o.put("texto", texto == null ? "" : texto);
            body = o.toString();
        } catch (Exception e) {
            if (alTerminar != null) alTerminar.run();
            return;
        }
        new Thread(() -> {
            boolean ok = post(url, body);
            if (ok) {
                log(c, canal + " " + app + ": enviado");
                reintentarCola(c, url);
            } else {
                encolar(c, body);
                log(c, canal + " " + app + ": sin conexión, se reintenta");
            }
            if (alTerminar != null) alTerminar.run();
        }).start();
    }

    static boolean post(String url, String body) {
        HttpURLConnection con = null;
        try {
            con = (HttpURLConnection) new URL(url).openConnection();
            con.setRequestMethod("POST");
            con.setDoOutput(true);
            con.setInstanceFollowRedirects(false);
            con.setConnectTimeout(15000);
            con.setReadTimeout(25000);
            con.setRequestProperty("Content-Type", "application/json; charset=utf-8");
            byte[] b = body.getBytes(StandardCharsets.UTF_8);
            try (OutputStream os = con.getOutputStream()) { os.write(b); }
            int code = con.getResponseCode();
            // Apps Script responde 302 (redirección) cuando ya procesó el pedido
            return code == 200 || code == 302 || code == 303;
        } catch (Exception e) {
            return false;
        } finally {
            if (con != null) con.disconnect();
        }
    }

    static synchronized void encolar(Context c, String body) {
        try {
            SharedPreferences p = prefs(c);
            JSONArray a = new JSONArray(p.getString("cola", "[]"));
            a.put(body);
            while (a.length() > 100) a.remove(0);
            p.edit().putString("cola", a.toString()).apply();
        } catch (Exception ignored) { }
    }

    static void reintentarCola(Context c, String url) {
        JSONArray a;
        synchronized (Sender.class) {
            try { a = new JSONArray(prefs(c).getString("cola", "[]")); } catch (Exception e) { return; }
            if (a.length() == 0) return;
            prefs(c).edit().putString("cola", "[]").apply();
        }
        for (int i = 0; i < a.length(); i++) {
            String b = a.optString(i);
            if (!post(url, b)) encolar(c, b);
        }
    }

    static synchronized void log(Context c, String linea) {
        SharedPreferences p = prefs(c);
        String hora = new SimpleDateFormat("dd/MM HH:mm", Locale.US).format(new Date());
        String prev = p.getString("log", "");
        String nuevo = hora + "  " + linea + "\n" + prev;
        String[] l = nuevo.split("\n");
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < Math.min(l.length, 15); i++) sb.append(l[i]).append('\n');
        p.edit().putString("log", sb.toString()).apply();
    }
}
