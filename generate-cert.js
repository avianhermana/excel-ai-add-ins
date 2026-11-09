const fs = require('fs');
const path = require('path');
const selfsigned = require('selfsigned');

async function generateCerts() {
  const certsDir = path.join(__dirname, 'certs');
  const certPath = path.join(certsDir, 'cert.pem');
  const keyPath = path.join(certsDir, 'key.pem');

  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    console.log('✓ Sertifikat SSL localhost sudah tersedia.');
    return;
  }

  if (!fs.existsSync(certsDir)) {
    fs.mkdirSync(certsDir, { recursive: true });
  }

  console.log('Membuat sertifikat SSL mandiri untuk localhost...');
  const attrs = [
    { name: 'commonName', value: 'localhost' },
    { name: 'countryName', value: 'ID' },
    { name: 'organizationName', value: 'Excel Copilot Local Dev' }
  ];

  const pems = await selfsigned.generate(attrs, {
    days: 365,
    keySize: 2048,
    algorithm: 'sha256'
  });

  fs.writeFileSync(certPath, pems.cert);
  fs.writeFileSync(keyPath, pems.private);
  console.log('✓ Sertifikat SSL localhost berhasil dibuat di folder certs/');
}

generateCerts().catch((err) => {
  console.error('Gagal membuat sertifikat:', err);
  process.exit(1);
});
