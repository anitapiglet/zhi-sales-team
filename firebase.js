/* ===== Firebase 初始化與雲端資料層 =====
   使用 CDN ES Module 匯入，不需要 npm/建置工具。
   ============================================================ */
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-app.js";
import {
  getFirestore, doc, onSnapshot, setDoc, getDoc
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-auth.js";

const firebaseConfig = {
  apiKey: "AIzaSyA06mP4U-lXt6rcgduG8SVf5Blz4NO46Pw",
  authDomain: "jenglish-cef0d.firebaseapp.com",
  projectId: "jenglish-cef0d",
  storageBucket: "jenglish-cef0d.firebasestorage.app",
  messagingSenderId: "1073493465502",
  appId: "1:1073493465502:web:477fa3c72f71fe2530de8f"
};

const fbApp = initializeApp(firebaseConfig);
const db = getFirestore(fbApp);
const auth = getAuth(fbApp);
const WORKSPACE_REF = doc(db, "workspaces", "default");

function setConnStatus(state, text){
  var el = document.getElementById("connStatus");
  if(!el) return;
  el.className = state;
  el.textContent = text;
}

window.__CLOUD = {
  ref: WORKSPACE_REF,
  onAuthChange: function(cb){ onAuthStateChanged(auth, cb); },
  login: function(email, pw){ return signInWithEmailAndPassword(auth, email, pw); },
  logout: function(){ return signOut(auth); },
  subscribe: function(onData){
    setConnStatus("", "連線中…");
    return onSnapshot(WORKSPACE_REF, function(snap){
      setConnStatus("ok", "☁️ 已同步");
      onData(snap.exists() ? snap.data() : null);
    }, function(err){
      setConnStatus("bad", "⚠️ 連線失敗：" + err.message);
      console.error("Firestore onSnapshot error", err);
    });
  },
  save: function(data){
    setConnStatus("", "儲存中…");
    return setDoc(WORKSPACE_REF, data).then(function(){
      setConnStatus("ok", "☁️ 已同步");
    }).catch(function(err){
      setConnStatus("bad", "⚠️ 儲存失敗：" + err.message);
      console.error("Firestore save error", err);
      if(window.__toastFallback) window.__toastFallback("雲端儲存失敗：" + err.message);
    });
  },
  ensureExists: function(blank){
    return getDoc(WORKSPACE_REF).then(function(snap){
      if(!snap.exists()) return setDoc(WORKSPACE_REF, blank);
    });
  }
};

// 登入畫面事件綁定（放在 module script 內，因為需要用到 login()）
document.addEventListener("DOMContentLoaded", function(){
  var btn = document.getElementById("loginBtn");
  var errEl = document.getElementById("loginErr");
  function doLogin(){
    errEl.textContent = "";
    var email = document.getElementById("loginEmail").value.trim();
    var pw = document.getElementById("loginPassword").value;
    if(!email || !pw){ errEl.textContent = "請輸入 Email 與密碼"; return; }
    btn.disabled = true; btn.textContent = "登入中…";
    window.__CLOUD.login(email, pw).catch(function(err){
      var msg = "登入失敗";
      if(err.code==="auth/invalid-credential" || err.code==="auth/wrong-password" || err.code==="auth/user-not-found") msg = "帳號或密碼錯誤";
      else if(err.code==="auth/too-many-requests") msg = "嘗試次數過多，請稍後再試";
      else if(err.code==="auth/configuration-not-found") msg = "管理員尚未啟用登入功能：請到 Firebase Console → Authentication → Sign-in method，啟用「電子郵件/密碼」";
      else msg = "登入失敗："+err.message;
      errEl.textContent = msg;
    }).finally(function(){
      btn.disabled = false; btn.textContent = "登入";
    });
  }
  btn.addEventListener("click", doLogin);
  document.getElementById("loginPassword").addEventListener("keydown", function(e){ if(e.key==="Enter") doLogin(); });
  document.getElementById("logoutBtn").addEventListener("click", function(){ window.__CLOUD.logout(); });

  window.__CLOUD.onAuthChange(function(user){
    if(user){
      document.getElementById("loginScreen").classList.add("hidden");
      document.getElementById("app").classList.remove("hidden");
      document.getElementById("currentUserLabel").textContent = user.email;
      if(window.__APP_START) window.__APP_START();
    } else {
      document.getElementById("loginScreen").classList.remove("hidden");
      document.getElementById("app").classList.add("hidden");
      if(window.__APP_STOP) window.__APP_STOP();
    }
  });
});
