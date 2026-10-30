// JScript bridge: JSON-lines RPC over stdin/stdout between the DSH host plugin
// and CorelDRAW X4. ASCII-only source so cscript never misreads the encoding.
(function () {
  var EXE = '';
  if (WScript.Arguments.Length > 0) EXE = String(WScript.Arguments(0));

  var PROGIDS = ['CorelDRAW.Application.14', 'CorelDRAW.Application', 'CorelDraw.Application'];

  function str(s) {
    var out = '"';
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c === 34) out += '\\"';
      else if (c === 92) out += '\\\\';
      else if (c === 10) out += '\\n';
      else if (c === 13) out += '\\r';
      else if (c === 9) out += '\\t';
      else if (c < 32 || c > 126) {
        var h = c.toString(16);
        out += '\\u' + '0000'.slice(h.length) + h;
      } else out += s.charAt(i);
    }
    return out + '"';
  }

  function stringify(v) {
    var t = typeof v;
    if (t === 'undefined' || t === 'function' || v === null) return 'null';
    if (t === 'boolean' || t === 'number') return String(v);
    if (t === 'string') return str(v);
    if (v instanceof Array) {
      var a = [];
      for (var i = 0; i < v.length; i++) a.push(stringify(v[i]));
      return '[' + a.join(',') + ']';
    }
    var o = [];
    for (var k in v) o.push(str(k) + ':' + stringify(v[k]));
    return '{' + o.join(',') + '}';
  }

  function parse(text) {
    return eval('(' + text + ')');
  }

  function errMsg(e) {
    return String((e && (e.message || e.description)) || e);
  }

  var app = null;
  var attachHow = '';
  var launchedPid = 0;

  function shell() {
    return new ActiveXObject('WScript.Shell');
  }

  function tryGet(prog) {
    try {
      app = GetObject('', prog);
      attachHow = 'running:' + prog;
      return true;
    } catch (e) {}
    return false;
  }

  function tryCreate(prog) {
    try {
      app = new ActiveXObject(prog);
      attachHow = 'launched:' + prog;
      return true;
    } catch (e) {}
    return false;
  }

  function launchExe(seconds) {
    if (!EXE) return false;
    try {
      var p = shell().Exec(EXE);
      if (p && p.ProcessID) launchedPid = p.ProcessID;
    } catch (e) {
      return false;
    }
    for (var i = 0; i < seconds * 2; i++) {
      WScript.Sleep(500);
      for (var j = 0; j < PROGIDS.length; j++) {
        if (tryGet(PROGIDS[j])) return true;
      }
    }
    return false;
  }

  // Attach to CorelDRAW. Probes running instances first, then a fresh COM
  // launch, then the green build's own exe (waiting up to 45s) if configured.
  function attach() {
    var i;
    if (app) return true;
    for (i = 0; i < PROGIDS.length; i++) {
      if (tryGet(PROGIDS[i])) return true;
    }
    for (i = 0; i < PROGIDS.length; i++) {
      if (tryCreate(PROGIDS[i])) return true;
    }
    return launchExe(45);
  }

  function runExec(code, args) {
    var fn;
    var body = 'var args = ' + stringify(args) + ';\n' + code;
    // eslint-disable-next-line no-new-func
    fn = new Function('app', 'WScript', 'ActiveXObject', 'GetObject', body);
    return fn(app, WScript, ActiveXObject, GetObject);
  }

  function doStatus(args) {
    var ok = attach();
    var res = {
      attached: ok,
      mode: ok ? 'com' : 'none',
      bridge: 'cscript/JScript',
      exePath: EXE,
      launchedPid: launchedPid,
      attach: attachHow,
      documents: 0,
      version: ''
    };
    if (ok) {
      try {
        res.version = String(app.Version);
      } catch (e) {
        res.version = '?';
      }
      try {
        res.documents = app.Documents.Count;
      } catch (e) {
        res.documents = -1;
      }
    }
    return res;
  }

  function doExec(args) {
    if (!attach()) {
      throw new Error('CorelDRAW is not attached (COM unavailable and exePath '
        + (EXE ? 'failed to launch ' + EXE : 'not configured') + ')');
    }
    return { value: runExec(args.code, args.args) };
  }

  function doOpen(args) {
    if (!attach()) throw new Error('CorelDRAW is not attached');
    var err = '';
    var d = null;
    try {
      d = app.OpenDocument(String(args.path));
    } catch (e1) {
      err += 'OpenDocument: ' + errMsg(e1) + '; ';
      try {
        d = app.CreateDocument();
        d.ActiveLayer.Import(String(args.path));
      } catch (e2) {
        throw new Error(err + 'Import: ' + errMsg(e2));
      }
    }
    return { name: d ? String(d.Name) : '' };
  }

  function doExport(args) {
    if (!attach()) throw new Error('CorelDRAW is not attached');
    var d = app.ActiveDocument;
    if (!d) throw new Error('no active document');
    var p = String(args.path);
    var ext = p.toLowerCase().replace(/^.*\./, '');
    var tried = [];
    if (args.filter === undefined || args.filter === null || args.filter === '') {
      if (ext === 'pdf') {
        try {
          d.PublishToPDF(p);
          tried.push('PublishToPDF ok');
        } catch (e) {
          tried.push('PublishToPDF: ' + errMsg(e));
        }
      }
      try {
        d.Export(p);
        tried.push('Export ok');
      } catch (e) {
        tried.push('Export: ' + errMsg(e));
      }
      if (tried.join('').indexOf(' ok') < 0) {
        try {
          d.SaveAs(p);
          tried.push('SaveAs ok');
        } catch (e) {
          tried.push('SaveAs: ' + errMsg(e));
        }
      }
    } else {
      try {
        d.Export(p, args.filter);
        tried.push('Export(filter=' + args.filter + ') ok');
      } catch (e) {
        tried.push('Export(filter=' + args.filter + '): ' + errMsg(e));
      }
    }
    return { document: String(d.Name), path: p, tried: tried };
  }

  function doUi(args) {
    var sh = shell();
    var target = args.pid ? args.pid : String(args.title || 'CorelDRAW');
    var activated = false;
    try {
      activated = sh.AppActivate(target);
    } catch (e) {
      activated = false;
    }
    if (activated && args.waitMs) WScript.Sleep(Number(args.waitMs));
    var sent = false;
    if (activated) {
      try {
        sh.SendKeys(String(args.keys));
        sent = true;
      } catch (e) {}
    }
    return { activated: activated, sent: sent, keys: String(args.keys), target: target };
  }

  function handle(req) {
    if (req.kind === 'status') return doStatus(req.args || {});
    if (req.kind === 'exec') return doExec(req.args || {});
    if (req.kind === 'open') return doOpen(req.args || {});
    if (req.kind === 'export') return doExport(req.args || {});
    if (req.kind === 'ui') return doUi(req.args || {});
    if (req.kind === 'quit') {
      app = null;
      return { bye: true };
    }
    throw new Error('unknown kind: ' + req.kind);
  }

  var stdin = WScript.StdIn;
  var stdout = WScript.StdOut;
  while (!stdin.AtEndOfStream) {
    var line = '';
    try {
      line = stdin.ReadLine();
    } catch (e) {
      break;
    }
    if (!line) continue;
    var res = { id: 0, ok: false, error: 'unhandled' };
    try {
      var req = parse(line);
      res.id = req.id;
      var out = handle(req);
      res.ok = true;
      res.result = out;
      delete res.error;
    } catch (e) {
      res.ok = false;
      res.error = errMsg(e);
    }
    stdout.WriteLine(stringify(res));
  }
})();
