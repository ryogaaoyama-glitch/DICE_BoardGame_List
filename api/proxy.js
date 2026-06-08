export default async function handler(req, res) {
  const params = new URLSearchParams(req.query).toString();
  const url = `https://script.google.com/macros/s/AKfycbxthxssVkUWIkWVaurUtAzI7MukRESiYPKaErDC4sEIYW-nnKfIBK-_IvbgSMOYAP_s/exec?${params}`;
  const response = await fetch(url, { redirect: 'follow' });
  const text = await response.text();
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Content-Type', 'application/json');
  res.status(200).send(text);
}
