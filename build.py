# Builds app.html from src/ (one self-contained file, so the update check covers everything)
import pathlib
s = pathlib.Path(__file__).parent / 'src'
t = (s / 'app.template.html').read_text(encoding='utf8')
for key, f in [('JSZIP', 'jszip.min.js'), ('ENGINE', 'engine.js'), ('APP', 'app.js')]:
    code = (s / f).read_text(encoding='utf8').replace('</script', '<\\/script')
    t = t.replace('/*' + key + '*/', code)
(s.parent / 'app.html').write_text(t, encoding='utf8')
print('app.html', len(t))
