"""Arma los archivos que se publican en GitHub (app/Index.html y app/Code.gs) a partir de las fuentes.

La pantalla se publica en dos partes:
- Index.html: estilos, estructura y un arranque chiquito (Google la procesa al servirla).
- El código de la pantalla va dentro de Code.gs como texto (PANTALLA_JS) y se pide con rpc('pantalla'),
  así no pasa por el procesamiento de Google, que alteraba el código.
"""
import json, re, sys, hashlib

src_html = open('Index.html', encoding='utf-8').read()
src_code = open('Code.gs', encoding='utf-8').read()

bloques = list(re.finditer(r'<script>(.*?)</script>', src_html, re.S))
main = bloques[-1]
js = main.group(1)

arranque = r'''<script>
(function(){
  var view=document.getElementById('view');
  function fallo(t){window.__errs=window.__errs||[];window.__errs.push(t);if(window.mostrarFallo)mostrarFallo();}
  if(!window.google||!google.script){fallo('Esta página tiene que abrirse desde la aplicación web de Google Apps Script.');return;}
  google.script.run.withSuccessHandler(function(r){
    try{
      var o=JSON.parse(r);
      if(!o.js||o.js.length!==o.largo){fallo('El código de la pantalla llegó incompleto ('+(o.js?o.js.length:0)+' de '+o.largo+'). En Apps Script ejecutá "instalar" otra vez.');return;}
      var s=document.createElement('script');s.textContent=o.js;document.body.appendChild(s);
    }catch(e){fallo('No pude arrancar la pantalla: '+e.message);}
  }).withFailureHandler(function(e){fallo('El servidor no respondió: '+(e&&e.message||e));}).rpc('pantalla','[]');
})();
</script>'''

html = src_html[:main.start()] + arranque + src_html[main.end():]

code = src_code.replace('var API_ = {', 'var API_ = {\n  pantalla: function () { return JSON.stringify({ js: PANTALLA_JS, largo: PANTALLA_JS.length }); },', 1)
code += '\n/* ---------- código de la pantalla (lo pide Index.html con rpc("pantalla")) ---------- */\nvar PANTALLA_JS = ' + json.dumps(js, ensure_ascii=True) + ';\n'

open(sys.argv[1] + '/Index.html', 'w', encoding='utf-8').write(html)
open(sys.argv[1] + '/Code.gs', 'w', encoding='utf-8').write(code)
print('Index.html', len(html), 'chars · Code.gs', len(code), 'chars · pantalla', len(js), 'chars')
