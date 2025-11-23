function createJwt_(clientEmail, privateKey) {
  const header = {
    alg: "RS256",
    typ: "JWT"
  };

  const now = Math.floor(Date.now() / 1000);
  const payload = {
    iss: clientEmail,
    sub: clientEmail,
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
    scope: "https://www.googleapis.com/auth/datastore"
  };

  const encode = obj => Utilities.base64EncodeWebSafe(JSON.stringify(obj));
  const signatureInput = `${encode(header)}.${encode(payload)}`;
  const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  return `${signatureInput}.${Utilities.base64EncodeWebSafe(signature)}`;
}

function getAccessToken_(jwt) {
  const tokenUrl = "https://oauth2.googleapis.com/token";
  const response = UrlFetchApp.fetch(tokenUrl, {
    method: "post",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    }
  });

  const result = JSON.parse(response.getContentText());
  return result.access_token;
}
