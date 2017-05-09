
  if (!String.prototype.format) {
    String.prototype.format = function() {
      var args = arguments;
      return this.replace(/{(\d+)}/g, function(match, number) { 
        return typeof args[number] != 'undefined'
          ? args[number]
          : match
        ;
      });
    };
  }

  if (!Array.prototype.contains) {
    Array.prototype.contains = function(x) {
      for (var i=0; i<this.length; i++) {
        if (this[i]==x) return true;
      }
      return false;
    };
  }

  if (!Array.prototype.remove) {
    Array.prototype.remove = function(x) {
      var result = [];
      for (var i=0; i<this.length; i++) {
        if (this[i]==x) continue;
        result.push(this[i]);
      }
      return result;
    };
  }

  if (!Array.prototype.removeAll) {
    Array.prototype.removeAll = function(a) {
      var result = [];
      for (var i=0; i<this.length; i++) {
        if (a.contains(this[i])) continue;
        result.push(this[i]);
      }
      return result;
    };
  }

  var __INTERNALS__ = {};

  var JSON = null;
  __INTERNALS__.json_html = new ActiveXObject('htmlfile');
  __INTERNALS__.json_html.write('<meta http-equiv="x-ua-compatible" content="IE=9" />');
  __INTERNALS__.json_html.close(JSON = __INTERNALS__.json_html.parentWindow.JSON);
  var FSO = new ActiveXObject("Scripting.FileSystemObject");
  var SHELL = new ActiveXObject("WScript.Shell");
  var SCRIPT_DIR = FSO.getParentFolderName(WScript.ScriptFullName);

  function make_msg_string(msg) {
    if (msg === undefined) return "undefined";
    if (typeof msg === 'object') return JSON.stringify(msg, null, 2);
    return msg;
  }

  function echo(msg) {
    //if (msg === undefined) msg = "undefined";
    WScript.Echo(make_msg_string(msg));
  }

  function msgbox(msg) {
    //if (msg === undefined) msg = "undefined";
    SHELL.Popup(make_msg_string(msg), 0, "Windows Script Host", 0);
  }

  function pause(msg) {
    if (msg === undefined) msg = "ポーズしています";
    echo(msg);
    msgbox(msg);
  }

  function write(s) {
    if (s === undefined) s = "";
    WScript.StdOut.Write(s);
  }

  function writeln(s) {
    if (s === undefined) s = "";
    WScript.StdOut.WriteLine(s);
  }

  function readTextFromFile_Utf8(path) {
    var StreamTypeEnum    = { adTypeBinary: 1, adTypeText: 2 };
    var SaveOptionsEnum   = { adSaveCreateNotExist: 1, adSaveCreateOverWrite: 2 };
    var LineSeparatorEnum = { adLF: 10, adCR: 13, adCRLF: -1 };
    var StreamReadEnum    = { adReadAll: -1, adReadLine: -2 };
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = StreamTypeEnum.adTypeText;
    stream.Charset = "utf-8";
    stream.LineSeparator = LineSeparatorEnum.adLF;
    stream.Open();
    stream.LoadFromFile(path);
    var result = stream.ReadText(StreamReadEnum.adReadAll);
    stream.Close();
    return result;
  }

  function readLinesFromFile_Utf8(path) {
    var StreamTypeEnum    = { adTypeBinary: 1, adTypeText: 2 };
    var SaveOptionsEnum   = { adSaveCreateNotExist: 1, adSaveCreateOverWrite: 2 };
    var LineSeparatorEnum = { adLF: 10, adCR: 13, adCRLF: -1 };
    var StreamReadEnum    = { adReadAll: -1, adReadLine: -2 };
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = StreamTypeEnum.adTypeText;
    stream.Charset = "utf-8";
    stream.LineSeparator = LineSeparatorEnum.adLF;
    var result = "";
    stream.Open();
    stream.LoadFromFile(path);
    var count = 0;
    while (!stream.EOS) {
      result += stream.ReadText(StreamReadEnum.adReadLine) + "\n";
      count++;
      if (count > 0 && (count % 100000) == 0) {
        WScript.Echo("readLinesFromFile_Utf8({0}): {1}".format(path, count));
      }
    }
    if (count >= 100000) WScript.Echo("readLinesFromFile_Utf8({0}): {1}...Done!".format(path, count));
    stream.Close();
    return result;
  }

  function writeTextToFile_Utf8(path, text, bom) {
    var StreamTypeEnum  = { adTypeBinary: 1, adTypeText: 2 };
    var SaveOptionsEnum = { adSaveCreateNotExist: 1, adSaveCreateOverWrite: 2 };
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = StreamTypeEnum.adTypeText;
    stream.Charset = "utf-8";
    stream.Open();
    stream.WriteText(text);
    if (!bom) {
      stream.Position = 0
      stream.Type = StreamTypeEnum.adTypeBinary;
      stream.Position = 3
      var buf = stream.Read();
      stream.Position = 0
      stream.Write(buf);
      stream.SetEOS();
    }
    stream.SaveToFile(path, SaveOptionsEnum.adSaveCreateOverWrite);
    stream.Close();
  }

  function readVectorFromFile_Utf8(path) {
    var json = readLinesFromFile_Utf8(path);
    var vec = JSON.parse(json);
    var result = [];
    for (var i=0; i<vec.length; i++) {
      result.push(vec[i]);
    }
    return result;
  }

  function readDictionaryFromFile_Utf8(path) {
    try {
      var json = readLinesFromFile_Utf8(path);
      return JSON.parse(json);
    } catch(e) {
      return {};
    }
  }

  function writeVectorToFile_Utf8(path, vec, bom) {
    var text = "[" + "\n";
    for (var i=0; i<vec.length; i++) {
      text += JSON.stringify(vec[i]);
      if (i<(vec.length-1)) text += ",";
      text += "\n";
    }
    text += "]";
    writeTextToFile_Utf8(path, text, bom);
  }


  function writeDictionaryToFile_Utf8(path, dict, bom) {
    writeTextToFile_Utf8(path, JSON.stringify(dict, null, 2), bom);
  }

  var SqliteDbClass = function(path) {

    var db = new ActiveXObject("ADODB.Connection");
    var driver = "DRIVER=SQLite3 ODBC Driver;Database={0}".format(path);

    this.Open = function() {
      db.Open(driver);
      return null;
    }

    this.Close = function() {
      db.Close();
      return null;
    }

    this.ExecuteUpdate = function(sql) {
      db.Execute(sql);
      return null;
    }

    this.ExecuteQuery = function(sql) {
      var rs = db.Execute(sql);
      var result = [];
      while (!rs.EOF) {
        var rec = {};
        for (var i=0; i<rs.Fields.Count; i++) {
          rec[rs.Fields(i).Name] = rs.Fields(i).Value;
        }
        result.push(rec);
        rs.MoveNext();
      }
      rs.Close();
      return result;
    }

  }

  var SqliteDictionaryClass = function(dictName) {

    var path = "{0}\\dict.{1}.db3".format(SCRIPT_DIR, dictName);
    var db = new SqliteDbClass(path);

    function create() {
      db.Open();
      var createSql = "create table if not exists dict (id integer primary key autoincrement, dict_key text unique, dict_value text);";
      db.ExecuteUpdate(createSql);
      db.Close();
    };

    function stringify(x) {
      if (x === undefined) return "undefined";
      return JSON.stringify(x);
    }

    function parse(x) {
      if (x === "undefined") return undefined;
      return JSON.parse(x);
    }

    this.contains = function(key) {
      var key_json = stringify(key);
      var selectSql = "select dict_key, dict_value from dict where dict_key=\"{0}\";".format(escape(key_json));
      db.Open();
      var rs = db.ExecuteQuery(selectSql);
      db.Close();
      return (rs.length > 0);
    };

    function escape(text) {
      return text.replace(new RegExp("\"", "g"), "\"\"");
    }

    this.put = function(key, val) {
      var key_json = stringify(key);
      var val_json = stringify(val);
      var sql;
      if (this.contains(key)) {
        sql = "update dict set dict_value=\"{0}\" where dict_key=\"{1}\";".format(escape(val_json), escape(key_json));
      } else {
        sql = "insert into dict (dict_key, dict_value) values (\"{0}\",\"{1}\");".format(escape(key_json), escape(val_json));
      }
      db.Open();
      db.ExecuteUpdate(sql);
      db.Close();
    };

    this.remove = function(key) {
      var key_json = stringify(key);
      var deleteSql = "delete from dict where dict_key=\"{0}\";".format(escape(key_json));
      db.Open();
      db.ExecuteUpdate(deleteSql);
      db.Close();
    };

    this.get = function(key) {
      var key_json = stringify(key);
      var selectSql = "select dict_key key, dict_value value from dict where dict_key=\"{0}\";".format(escape(key_json));
      db.Open();
      var rs = db.ExecuteQuery(selectSql);
      var result = undefined;
      if (rs.length > 0) result = parse(rs[0].value);
      db.Close();
      return result;
    };

    this.clear = function() {
      var sql = "delete from dict;";
      db.Open();
      db.ExecuteUpdate(sql);
      db.Close();
    };

    this.records = function(key) {
      var key_json = stringify(key);
      var selectSql = "select dict_key key, dict_value value from dict order by id";
      db.Open();
      var rs = db.ExecuteQuery(selectSql);
      var created = new Date().getTime();
      var result = [];
      for (var i=0; i<rs.length; i++) {
        var rec = rs[i];
        rec.key = parse(rec.key);
        rec.value = parse(rec.value);
        if (!rec.hasOwnProperty("created")) rec.created = created;
        if (!rec.hasOwnProperty("updated")) rec.updated = 0;
        delete rec.hasOwnProperty;
        result.push(rec);
      }
      db.Close();
      return result;
    };

    create();

  };

