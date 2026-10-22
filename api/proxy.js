export default async function handler(req, res) {
  const params = new URLSearchParams(req.query).toString();
  const url = `https://script.google.com/macros/s/AKfycbzjCLUOJfa6t-fNtoRFjyxEnppZ5L3VvSQ0_vRhUeOBWt399cPMLm_lMp7rj0sLi3qN/exec?${params}`;
  const response = await fetch(url, { redirect: 'follow' });
  const text = await response.text();
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Content-Type', 'application/json');
  res.status(200).send(text);
}
