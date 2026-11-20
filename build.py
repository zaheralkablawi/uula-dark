# Builds app.html from src/ (one self-contained file, so the update check covers everything)
import pathlib
s = pathlib.Path(__file__).parent / 'src'
t = (s / 'app.template.html').read_text(encoding='utf8')
# standalone UULA Dark: both parts shown (UULA Toolkit sets the mode to 'dark' or 'cover' when it opens this page)
for key, val in [('MODE', ''), ('TITLE', 'UULA Dark'), ('SUB', 'Dark')]:
    t = t.replace('/*' + key + '*/', val)
for key, f in [('I18N', 'i18n.js'), ('JSZIP', 'jszip.min.js'), ('ENGINE', 'engine.js'), ('APP', 'app.js')]:
    code = (s / f).read_text(encoding='utf8').replace('</script', '<\\/script')
    t = t.replace('/*' + key + '*/', code)
(s.parent / 'app.html').write_text(t, encoding='utf8')
print('app.html', len(t))
