const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');

const PORT = 3001;

const requestHandler = (req, res) => {
  if (req.method === 'POST' && req.url === '/api/claude') {
    let body = '';
    
    req.on('data', chunk => {
      body += chunk.toString();
    });
    
    req.on('end', () => {
      try {
        const requestData = JSON.parse(body);
        
        if (!requestData.apiKey) {
          res.writeHead(400, {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ error: 'API key is required' }));
          return;
        }
        
        const options = {
          hostname: 'api.anthropic.com',
          path: '/v1/messages',
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': requestData.apiKey,
            'anthropic-version': '2023-06-01'
          }
        };
        
        // Build request body - support system messages if present
        const apiRequestBody = {
          model: requestData.model,
          max_tokens: requestData.max_tokens,
          messages: requestData.messages
        };
        
        // Extract system message if first message has role 'system'
        // Claude API expects system messages in a separate 'system' field
        if (requestData.messages && requestData.messages.length > 0 && requestData.messages[0].role === 'system') {
          apiRequestBody.system = requestData.messages[0].content;
          // Remove system message from messages array (Claude API expects it separately)
          apiRequestBody.messages = requestData.messages.slice(1);
        }
        
        const apiReq = https.request(options, (apiRes) => {
          let apiBody = '';
          
          apiRes.on('data', (chunk) => {
            apiBody += chunk;
          });
          
          apiRes.on('end', () => {
            res.writeHead(apiRes.statusCode, {
              'Content-Type': 'application/json',
              'Access-Control-Allow-Origin': '*',
              'Access-Control-Allow-Methods': 'POST, OPTIONS',
              'Access-Control-Allow-Headers': 'Content-Type'
            });
            res.end(apiBody);
          });
        });
        
        apiReq.on('error', (error) => {
          res.writeHead(500, {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          });
          res.end(JSON.stringify({ error: error.message }));
        });
        
        apiReq.write(JSON.stringify(apiRequestBody));
        apiReq.end();
      } catch (error) {
        res.writeHead(400, {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        });
        res.end(JSON.stringify({ error: 'Invalid request: ' + error.message }));
      }
    });
  } else if (req.method === 'OPTIONS') {
    res.writeHead(200, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    res.end();
  } else {
    res.writeHead(404, {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Origin': '*'
    });
    res.end(JSON.stringify({ error: 'Not found' }));
  }
};

try {
  const certPath = path.join(process.env.HOME || process.env.USERPROFILE || '', '.office-addin-dev-certs', 'localhost.crt');
  const keyPath = path.join(process.env.HOME || process.env.USERPROFILE || '', '.office-addin-dev-certs', 'localhost.key');
  
  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    const options = {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath)
    };
    
    const server = https.createServer(options, requestHandler);
    server.listen(PORT, () => {
      console.log(`Claude API proxy server running on https://localhost:${PORT}`);
      console.log('Keep this running while using the WordTrack add-in');
    });
  } else {
    console.log('HTTPS certificates not found, using HTTP (may have mixed content issues)');
    console.log('To fix: Run "npx office-addin-dev-certs install" first');
    const server = http.createServer(requestHandler);
    server.listen(PORT, () => {
      console.log(`Claude API proxy server running on http://localhost:${PORT}`);
      console.log('Keep this running while using the WordTrack add-in');
    });
  }
} catch (error) {
  console.error('Error starting server:', error);
  console.log('Falling back to HTTP server');
  const server = http.createServer(requestHandler);
  server.listen(PORT, () => {
    console.log(`Claude API proxy server running on http://localhost:${PORT}`);
    console.log('Keep this running while using the WordTrack add-in');
  });
}
