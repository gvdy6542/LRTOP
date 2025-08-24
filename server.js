const http = require('http');
const crypto = require('crypto');

const users = {
  admin: '8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918'
};

const sessions = new Map();

function handleLogin(req, res) {
  let body = '';
  req.on('data', chunk => body += chunk);
  req.on('end', () => {
    try {
      const { username, password } = JSON.parse(body || '{}');
      const storedHash = users[username];
      if (!storedHash) {
        res.writeHead(401, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ message: 'Invalid username or password' }));
        return;
      }
      const hash = crypto.createHash('sha256').update(password).digest('hex');
      if (hash !== storedHash) {
        res.writeHead(401, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ message: 'Invalid username or password' }));
        return;
      }
      const token = crypto.randomBytes(24).toString('hex');
      sessions.set(token, username);
      res.writeHead(200, {
        'Content-Type': 'application/json',
        'Set-Cookie': `token=${token}; HttpOnly`
      });
      res.end(JSON.stringify({ token }));
    } catch (err) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ message: 'Bad request' }));
    }
  });
}

const server = http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/login') {
    handleLogin(req, res);
  } else {
    res.writeHead(404);
    res.end();
  }
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
