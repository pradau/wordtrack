# CORS Proxy Setup for Claude API

Office Add-ins run in a browser sandbox that blocks direct API calls to external services due to CORS (Cross-Origin Resource Sharing) restrictions. To use the Claude API, you need a proxy server.

## Option 1: Simple Node.js Proxy Server (Recommended)

Create a simple proxy server that forwards requests to the Claude API.

### Step 1: Create Proxy Server File

Create a new file `proxy-server.js` in the project root:

```javascript
const http = require('http');
const https = require('https');

const PORT = 3001;

const server = http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/api/claude') {
    let body = '';
    
    req.on('data', chunk => {
      body += chunk.toString();
    });
    
    req.on('end', () => {
      const requestData = JSON.parse(body);
      
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
      
      const requestBody = {
        model: requestData.model,
        max_tokens: requestData.max_tokens,
        messages: requestData.messages
      };
      
      apiReq.write(JSON.stringify(requestBody));
      apiReq.end();
    });
  } else if (req.method === 'OPTIONS') {
    res.writeHead(200, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    res.end();
  } else {
    res.writeHead(404);
    res.end('Not found');
  }
});

server.listen(PORT, () => {
  console.log(`Proxy server running on http://localhost:${PORT}`);
});
```

### Step 2: Update package.json

Add a script to run the proxy:

```json
"scripts": {
  "proxy": "node proxy-server.js"
}
```

### Step 3: Update the Add-in Code

The add-in code needs to be updated to use the proxy. This will be done in the next update.

### Step 4: Run the Proxy

In a separate terminal:

```bash
npm run proxy
```

Keep this running while using the add-in.

## Option 2: Use a Public CORS Proxy (Not Recommended for Production)

For quick testing, you could use a public CORS proxy, but this is **not secure** as it exposes your API key. Only use for development.

## Security Note

The proxy server approach keeps your API key server-side, which is more secure than using a public CORS proxy. However, for personal use, the current localStorage approach is acceptable.

