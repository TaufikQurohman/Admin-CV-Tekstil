/* ============================================================
   LOGIN / SESSION
   ============================================================ */

function login(username, password) {

  const user =
    cleanString_(username);

  const pass =
    String(
      password == null
        ? ""
        : password
    );


  if (
    user !== "admin" ||
    pass !== "Bismillah123"
  ) {

    throw new Error(
      "Username atau password salah."
    );

  }


  const token =
    Utilities.getUuid() +
    "-" +
    Utilities.getUuid();


  CacheService
    .getScriptCache()
    .put(
      "AUTH_" + token,
      "admin",
      CONFIG.SESSION_SECONDS
    );


  return {
    success: true,
    token: token,
    username: "admin"
  };

}

function checkLogin(token) {

  if (!token) {

    return {
      loggedIn: false,
      username: ""
    };

  }


  const username =
    CacheService
      .getScriptCache()
      .get(
        "AUTH_" + String(token)
      );


  return {
    loggedIn: username === "admin",
    username:
      username || ""
  };

}

function logout(token) {

  if (token) {

    CacheService
      .getScriptCache()
      .remove(
        "AUTH_" + String(token)
      );

  }


  return {
    success: true
  };

}

function requireAuth_(token) {

  if (!token) {

    throw new Error(
      "Sesi login tidak ditemukan. Silakan login kembali."
    );

  }


  const username =
    CacheService
      .getScriptCache()
      .get(
        "AUTH_" + String(token)
      );


  if (username !== "admin") {

    throw new Error(
      "Sesi login tidak valid atau sudah kedaluwarsa. Silakan login kembali."
    );

  }


  return username;

}


