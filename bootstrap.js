// ================================================================
//  bootstrap.js: 文字化け本体を実行時に元へ戻して読み込む起動部  【大A / 中0 / 小0-0】
//  ①修正用SSに置く。ここと blob.js（文字化けした本体）だけがGAS上に上がり、
//  読める本体（コード.js）はローカルとgitにだけ残す（客・閲覧者にはコードが見えない）。
//  戻す鍵は①のスクリプトプロパティ DECKEY_ にだけ置く（閲覧権限では見えない）。
//  速さ対策：一度戻した本体を CacheService に分割保存し、次回以降は戻し処理を省く。
//  ライブラリとして呼ばれた時も、①自身のWebアプリ・メニューでも、読み込み時に一度だけ動く。
// ================================================================
(function () {
  try {
    var ver = (typeof CODEVER_ !== 'undefined') ? CODEVER_ : '0'; // 反映のたびに変わる版番号でキャッシュを分ける（古いコードを返さない）
    var cache = CacheService.getScriptCache();
    var src = null;
    var nStr = cache.get('SRC_' + ver + '_N');
    if (nStr) {
      var n = parseInt(nStr, 10), keys = [];
      for (var i = 0; i < n; i++) keys.push('SRC_' + ver + '_' + i);
      var got = cache.getAll(keys), parts = [], ok = true;
      for (var j = 0; j < n; j++) { var v = got['SRC_' + ver + '_' + j]; if (v == null) { ok = false; break; } parts.push(v); }
      if (ok) src = parts.join('');
    }
    if (src == null) {
      var key = PropertiesService.getScriptProperties().getProperty('DECKEY_');
      if (!key) return; // 鍵未設定なら何もしない（動かない状態のまま止める）
      var b = Utilities.base64Decode(BLOB_), k = Utilities.base64Decode(key);
      for (var m = 0; m < b.length; m++) b[m] = b[m] ^ k[m % 32];
      src = Utilities.newBlob(b).getDataAsString('UTF-8');
      var put = {}, cnt = 0, CH = 90000;
      for (var p = 0; p < src.length; p += CH) { put['SRC_' + ver + '_' + cnt] = src.substr(p, CH); cnt++; }
      cache.putAll(put, 21600);
      cache.put('SRC_' + ver + '_N', String(cnt), 21600);
    }
    (0, eval)(src);
  } catch (e) {
    try { console.error('bootstrap失敗: ' + e); } catch (e2) {}
  }
})();
