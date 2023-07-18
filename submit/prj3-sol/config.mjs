const CERT_BASE = `${process.env.HOME}/tmp/localhost-certs`;

export default {

  auth: {
    dbUrl:  'mongodb://localhost:27017/ss',
  },

  ws: {
    port: 2345,
  },

  https: {
    certPath: `${CERT_BASE}/localhost.crt`,
    keyPath: `${CERT_BASE}/localhost.key`,
  },
  

};
