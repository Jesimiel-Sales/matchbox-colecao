"""Junta Index.html + Styles.html + App.html num único Index.html (o que vai para o Apps Script)."""
import pathlib
src = pathlib.Path(__file__).parent / 'src'
idx = (src / 'Index.html').read_text(encoding='utf-8')
idx = idx.replace("<?!= include('Styles'); ?>", (src / 'Styles.html').read_text(encoding='utf-8'))
idx = idx.replace("<?!= include('App'); ?>", (src / 'App.html').read_text(encoding='utf-8'))
assert '<?' not in idx
pathlib.Path(__file__).parent.joinpath('dist').mkdir(exist_ok=True)
(pathlib.Path(__file__).parent / 'dist' / 'Index.html').write_text(idx, encoding='utf-8')
print('dist/Index.html:', len(idx), 'caracteres')
