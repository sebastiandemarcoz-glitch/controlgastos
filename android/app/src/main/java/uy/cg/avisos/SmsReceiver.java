package uy.cg.avisos;

import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.provider.Telephony;
import android.telephony.SmsMessage;

/** Lee los SMS que llegan y manda los que parecen movimientos de dinero. */
public class SmsReceiver extends BroadcastReceiver {
    @Override
    public void onReceive(Context c, Intent i) {
        if (!Telephony.Sms.Intents.SMS_RECEIVED_ACTION.equals(i.getAction())) return;
        SmsMessage[] msgs = Telephony.Sms.Intents.getMessagesFromIntent(i);
        if (msgs == null || msgs.length == 0) return;
        StringBuilder sb = new StringBuilder();
        String de = "";
        for (SmsMessage m : msgs) {
            if (m == null) continue;
            if (m.getDisplayOriginatingAddress() != null) de = m.getDisplayOriginatingAddress();
            if (m.getDisplayMessageBody() != null) sb.append(m.getDisplayMessageBody());
        }
        String texto = sb.toString();
        if (!Sender.pareceMovimiento(texto)) return;
        final PendingResult pr = goAsync();
        Sender.enviar(c, "SMS", de, "", texto, pr::finish);
    }
}
