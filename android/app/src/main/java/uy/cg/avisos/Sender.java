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
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.List;

/** Manda cada aviso de movimiento al receptor de la app de gastos (Google Apps Script). */
public final class Sender {
    private Sender() {}

    /** Palabras clave de fábrica: alcanza con que el mensaje tenga una (sin importar mayúsculas ni tildes). */
    static final String CLAVES_FABRICA =
            "compra\nconsumo\ndébito\ntransacción\ntransferencia\nacreditación\nacreditó\npago\nretiro\ncobro\ndepósito\ncargo\noperación\n$\nu$s\nusd\nuyu";

    static String normalizar(String t) {
        String s = Normalizer.normalize(t == null ? "" : t, Normalizer.Form.NFD).replaceAll("\\p{M}+", "");
        return s.toLowerCase(Locale.ROOT).replaceAll("\\s+", " ");
    }

    static String clavesTexto(Context c) {
        return prefs(c).getString("claves", CLAVES_FABRICA);
    }

    static List<String> claves(Context c) {
        List<String> r = new ArrayList<>();
        for (String w : clavesTexto(c).split("[\\n,;]+")) {
            String n = normalizar(w).trim();
            if (!n.isEmpty()) r.add(n);
        }
        return r;
    }

    /** Devuelve la palabra clave que aparece en el mensaje, o null si no tiene ninguna (y entonces no se manda). */
    static String palabraEncontrada(Context c, String t) {
        String n = normalizar(t);
        if (n.trim().isEmpty()) return null;
        for (String w : claves(c)) if (n.contains(w)) return w;
        return null;
    }

    static boolean pareceMovimiento(Context c, String t) {
        return palabraEncontrada(c, t) != null;
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
            int r = post(url, body);
            if (r == OK) {
                log(c, canal + " " + app + ": enviado ✓");
                reintentarCola(c, url);
            } else if (r == PIDE_LOGIN) {
                log(c, "✗ La dirección pide iniciar sesión: usá la implementación con acceso \"Cualquier persona\"");
            } else if (r == RECHAZADO) {
                log(c, "✗ El receptor no aceptó el aviso (clave o dirección incorrecta)");
            } else {
                encolar(c, body);
                log(c, canal + " " + app + ": sin conexión, se reintenta");
            }
            if (alTerminar != null) alTerminar.run();
        }).start();
    }

    static final int OK = 0, SIN_RED = 1, PIDE_LOGIN = 2, RECHAZADO = 3;

    static int post(String url, String body) {
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
            if (code == 302 || code == 303 || code == 301) {
                // Apps Script responde con una redirección a googleusercontent cuando ya guardó el aviso;
                // si redirige a accounts.google.com es porque la dirección pide iniciar sesión.
                String loc = con.getHeaderField("Location");
                if (loc != null && (loc.contains("accounts.google.com") || loc.contains("ServiceLogin"))) return PIDE_LOGIN;
                if (loc != null && loc.contains("googleusercontent.com")) {
                    String resp = leer(loc);
                    if (resp != null && resp.contains("clave incorrecta")) return RECHAZADO;
                    return OK;
                }
                return RECHAZADO;
            }
            if (code == 200) return OK;
            if (code == 401 || code == 403) return PIDE_LOGIN;
            return code >= 500 ? SIN_RED : RECHAZADO;
        } catch (Exception e) {
            return SIN_RED;
        } finally {
            if (con != null) con.disconnect();
        }
    }

    static String leer(String url) {
        HttpURLConnection con = null;
        try {
            con = (HttpURLConnection) new URL(url).openConnection();
            con.setConnectTimeout(15000);
            con.setReadTimeout(20000);
            java.io.InputStream in = con.getInputStream();
            byte[] buf = new byte[512];
            int n = in.read(buf);
            return n > 0 ? new String(buf, 0, n, StandardCharsets.UTF_8) : "";
        } catch (Exception e) {
            return null;
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
            if (post(url, b) == SIN_RED) encolar(c, b);
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
