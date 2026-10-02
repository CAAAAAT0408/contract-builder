// 브라우저 안에서 Python(Pyodide)을 띄워 계약서 엔진을 돌린다.
// 계약서·공고 파일은 이 브라우저 밖으로 나가지 않는다.
const VERSION = '2026.10.02';
const PYODIDE = 'https://cdn.jsdelivr.net/pyodide/v0.29.5/full/';
const PY_FILES = ['hwp_reader.py', 'notice.py', 'contract.py', 'engine.py', 'report.py', 'bridge.py'];
const WHEELS = ['olefile-0.47-py2.py3-none-any.whl', 'et_xmlfile-2.0.0-py3-none-any.whl', 'openpyxl-3.1.5-py2.py3-none-any.whl'];

importScripts(PYODIDE + 'pyodide.js');

let pyodide = null;
const status = msg => postMessage({status: msg});

async function fetchBytes(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`${url} 을(를) 받지 못했습니다 (${r.status})`);
  return new Uint8Array(await r.arrayBuffer());
}

async function init() {
  status('Python 준비 중… (처음 한 번은 10~20초)');
  pyodide = await loadPyodide({indexURL: PYODIDE});
  status('부품 받는 중…');
  await pyodide.loadPackage(['lxml']);
  pyodide.FS.mkdirTree('/home/pyodide/app');
  pyodide.FS.mkdirTree('/home/pyodide/wheels');
  for (const w of WHEELS) pyodide.FS.writeFile('/home/pyodide/wheels/' + w, await fetchBytes(`py/wheels/${w}`));
  for (const f of PY_FILES) pyodide.FS.writeFile('/home/pyodide/app/' + f, await fetchBytes(`py/${f}?v=${VERSION}`));
  await pyodide.runPythonAsync(`
import sys, zipfile, glob
for w in glob.glob('/home/pyodide/wheels/*.whl'):
    zipfile.ZipFile(w).extractall('/home/pyodide/lib')
sys.path[:0] = ['/home/pyodide/app', '/home/pyodide/lib']
import bridge
`);
  status('');
  postMessage({ready: true});
}
const ready = init().catch(e => { postMessage({fatal: String(e && e.message || e)}); throw e; });

function call(fn, ...args) {
  const bridge = pyodide.globals.get('bridge');
  try {
    const f = bridge[fn];
    const out = f(...args);
    f.destroy && f.destroy();
    return out;
  } finally { bridge.destroy(); }
}

onmessage = async (e) => {
  const {id, cmd, args} = e.data;
  try {
    await ready;
    let result;
    if (cmd === 'analyze') {
      result = JSON.parse(call('analyze', args.c, args.n, args.cname, args.nname, args.strike));
    } else if (cmd === 'reanalyze') {
      result = JSON.parse(call('reanalyze', args.strike));
    } else if (cmd === 'build') {
      result = JSON.parse(call('build', JSON.stringify(args)));
      if (result.ok) {
        for (const k of ['docx', 'xlsx']) {
          const py = call('result', k);
          result[k] = py.toJs();
          py.destroy();
        }
      }
    } else {
      throw new Error('알 수 없는 명령: ' + cmd);
    }
    postMessage({id, result}, result && result.docx ? [result.docx.buffer, result.xlsx.buffer] : []);
  } catch (err) {
    postMessage({id, result: {ok: false, error: '처리 중 오류가 났습니다: ' + String(err && err.message || err).split('\n').slice(-3).join(' ')}});
  }
};
