package uy.cg.avisos;

import android.Manifest;
import android.app.Activity;
import android.content.ComponentName;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.graphics.Color;
import android.graphics.Typeface;
import android.net.Uri;
import android.os.Bundle;
import android.provider.Settings;
import android.text.InputType;
import android.view.Gravity;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.TextView;
import android.widget.Toast;

/** Pantalla única: permisos, configuración (dirección y clave) y últimos envíos. */
public class MainActivity extends Activity {
    private TextView estadoNotif, estadoSms, estadoCfg, log;
    private EditText url, clave;

    @Override
    protected void onCreate(Bundle b) {
        super.onCreate(b);
        int pad = dp(18);
        LinearLayout l = new LinearLayout(this);
        l.setOrientation(LinearLayout.VERTICAL);
        l.setPadding(pad, pad, pad, pad);

        TextView t = texto("Gastos · Avisos", 22, true);
        l.addView(t);
        l.addView(texto("Manda a tu app de gastos los avisos de compras que llegan por notificación o SMS. "
                + "Solo se mandan los mensajes que traen un importe ($ o U$S); el resto no sale del celular.", 14, false));

        l.addView(titulo("1. Conectar con tu app de gastos"));
        estadoCfg = texto("", 14, false);
        l.addView(estadoCfg);
        l.addView(texto("Lo más fácil: en la app de gastos andá a Ajustes → Avisos y tocá \"Configurar la app de avisos\". "
                + "O pegá acá la dirección y la clave:", 13, false));
        url = campo("Dirección del receptor (https://script.google.com/…/exec)");
        clave = campo("Clave");
        l.addView(url);
        l.addView(clave);
        l.addView(boton("Guardar", v -> guardar(url.getText().toString(), clave.getText().toString())));

        l.addView(titulo("2. Permitir leer notificaciones"));
        estadoNotif = texto("", 14, false);
        l.addView(estadoNotif);
        l.addView(boton("Abrir permiso de notificaciones", v ->
                startActivity(new Intent(Settings.ACTION_NOTIFICATION_LISTENER_SETTINGS))));

        l.addView(titulo("3. Permitir leer SMS"));
        estadoSms = texto("", 14, false);
        l.addView(estadoSms);
        l.addView(boton("Permitir SMS", v -> requestPermissions(new String[]{Manifest.permission.RECEIVE_SMS}, 1)));

        l.addView(texto("Si Android dice \"Configuración restringida\" o no te deja activar un permiso: tocá el botón de abajo, "
                + "en la pantalla de la app abrí el menú ⋮ (arriba a la derecha) → \"Permitir configuración restringida\" y volvé a intentar.", 13, false));
        l.addView(boton("Abrir información de la app", v -> startActivity(new Intent(Settings.ACTION_APPLICATION_DETAILS_SETTINGS,
                Uri.parse("package:" + getPackageName())))));

        l.addView(titulo("Probar"));
        l.addView(boton("Mandar un aviso de prueba", v -> {
            Sender.enviar(this, "APP", "Prueba", "Aviso de prueba",
                    "Compra por $ 1,00 en PRUEBA GASTOS AVISOS", () -> runOnUiThread(this::refrescar));
            Toast.makeText(this, "Enviado: aparece en \"Clasificar\" en unos minutos (descartalo)", Toast.LENGTH_LONG).show();
        }));
        l.addView(titulo("Últimos envíos"));
        log = texto("", 12, false);
        log.setTypeface(Typeface.MONOSPACE);
        l.addView(log);

        ScrollView sv = new ScrollView(this);
        sv.addView(l);
        setContentView(sv);
        leerEnlace(getIntent());
    }

    @Override
    protected void onNewIntent(Intent i) {
        super.onNewIntent(i);
        leerEnlace(i);
    }

    @Override
    protected void onResume() {
        super.onResume();
        refrescar();
    }

    @Override
    public void onRequestPermissionsResult(int rc, String[] p, int[] r) {
        super.onRequestPermissionsResult(rc, p, r);
        refrescar();
    }

    private void leerEnlace(Intent i) {
        Uri d = i == null ? null : i.getData();
        if (d == null || !"cgavisos".equals(d.getScheme())) return;
        String u = d.getQueryParameter("u"), k = d.getQueryParameter("k");
        if (u != null && k != null) guardar(u, k);
    }

    private void guardar(String u, String k) {
        u = u.trim();
        k = k.trim();
        if (!u.startsWith("https://") || k.isEmpty()) {
            Toast.makeText(this, "Revisá la dirección y la clave", Toast.LENGTH_LONG).show();
            return;
        }
        Sender.prefs(this).edit().putString("url", u).putString("k", k).apply();
        Toast.makeText(this, "Conectada con tu app de gastos", Toast.LENGTH_SHORT).show();
        refrescar();
    }

    private void refrescar() {
        SharedPreferences p = Sender.prefs(this);
        String u = p.getString("url", ""), k = p.getString("k", "");
        if (url.getText().length() == 0) url.setText(u);
        if (clave.getText().length() == 0) clave.setText(k);
        marcar(estadoCfg, !u.isEmpty() && !k.isEmpty(), "Conectada", "Falta conectar");
        String hab = Settings.Secure.getString(getContentResolver(), "enabled_notification_listeners");
        boolean notif = hab != null && hab.contains(new ComponentName(this, NotifListener.class).flattenToString());
        marcar(estadoNotif, notif, "Activado", "Falta activar: buscá \"Gastos · Avisos\" en la lista y activalo");
        boolean sms = checkSelfPermission(Manifest.permission.RECEIVE_SMS) == PackageManager.PERMISSION_GRANTED;
        marcar(estadoSms, sms, "Activado", "Falta permitir");
        log.setText(p.getString("log", "Todavía no se mandó nada."));
    }

    private void marcar(TextView v, boolean ok, String si, String no) {
        v.setText(ok ? "✓ " + si : "✗ " + no);
        v.setTextColor(ok ? Color.rgb(20, 145, 95) : Color.rgb(210, 60, 60));
    }

    private int dp(int v) { return (int) (v * getResources().getDisplayMetrics().density); }

    private TextView texto(String s, int sp, boolean negrita) {
        TextView t = new TextView(this);
        t.setText(s);
        t.setTextSize(sp);
        if (negrita) t.setTypeface(Typeface.DEFAULT_BOLD);
        t.setPadding(0, dp(4), 0, dp(4));
        return t;
    }

    private TextView titulo(String s) {
        TextView t = texto(s, 17, true);
        t.setPadding(0, dp(20), 0, dp(4));
        return t;
    }

    private EditText campo(String hint) {
        EditText e = new EditText(this);
        e.setHint(hint);
        e.setSingleLine(true);
        e.setInputType(InputType.TYPE_CLASS_TEXT | InputType.TYPE_TEXT_VARIATION_URI);
        return e;
    }

    private Button boton(String s, View.OnClickListener l) {
        Button b = new Button(this);
        b.setText(s);
        b.setAllCaps(false);
        b.setGravity(Gravity.CENTER);
        b.setOnClickListener(l);
        return b;
    }
}
