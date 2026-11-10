package uy.cg.avisos;

import android.app.Notification;
import android.content.pm.ApplicationInfo;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.service.notification.NotificationListenerService;
import android.service.notification.StatusBarNotification;

import java.util.Arrays;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

/** Lee las notificaciones de las apps (bancos) y manda las que parecen movimientos de dinero. */
public class NotifListener extends NotificationListenerService {

    // Los SMS los lee SmsReceiver y los mails los lee la app de Google: se excluyen acá para no mandarlos dos veces.
    static final Set<String> EXCLUIR = new HashSet<>(Arrays.asList(
            "com.google.android.apps.messaging", "com.samsung.android.messaging", "com.android.mms",
            "com.google.android.gm", "com.microsoft.office.outlook", "com.samsung.android.email.provider",
            "com.whatsapp", "com.whatsapp.w4b", "org.telegram.messenger", "com.facebook.orca",
            "com.instagram.android", "com.android.systemui"));

    private final Map<String, Long> recientes = new LinkedHashMap<String, Long>() {
        @Override protected boolean removeEldestEntry(Map.Entry<String, Long> e) { return size() > 60; }
    };

    @Override
    public void onNotificationPosted(StatusBarNotification sbn) {
        try {
            String pkg = sbn.getPackageName();
            if (pkg.equals(getPackageName()) || EXCLUIR.contains(pkg) || sbn.isOngoing()) return;
            Notification n = sbn.getNotification();
            if ((n.flags & Notification.FLAG_GROUP_SUMMARY) != 0) return;
            Bundle e = n.extras;
            CharSequence t = e.getCharSequence(Notification.EXTRA_TITLE);
            CharSequence x = e.getCharSequence(Notification.EXTRA_TEXT);
            CharSequence b = e.getCharSequence(Notification.EXTRA_BIG_TEXT);
            String titulo = t == null ? "" : t.toString();
            String texto = b != null && b.length() > 0 ? b.toString() : (x == null ? "" : x.toString());
            if (!Sender.pareceMovimiento(this, titulo + " " + texto)) return;
            String clave = pkg + "|" + titulo + "|" + texto;
            long ahora = System.currentTimeMillis();
            synchronized (recientes) {
                Long antes = recientes.get(clave);
                if (antes != null && ahora - antes < 10 * 60 * 1000) return;   // la misma notificación actualizada
                recientes.put(clave, ahora);
            }
            Sender.enviar(this, "APP", nombreApp(pkg), titulo, texto, null);
        } catch (Exception ignored) { }
    }

    private String nombreApp(String pkg) {
        try {
            PackageManager pm = getPackageManager();
            ApplicationInfo ai = pm.getApplicationInfo(pkg, 0);
            return pm.getApplicationLabel(ai).toString();
        } catch (Exception e) {
            return pkg;
        }
    }
}
