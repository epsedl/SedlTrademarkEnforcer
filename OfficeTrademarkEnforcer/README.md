# Office Trademark Enforcer Add-in

This is an Office Add-in for trademark enforcement.

## Deployment on Ubuntu

1. Clone the repository:
   ```
   git clone <your-github-repo-url>
   cd OfficeTrademarkEnforcer
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. To run the production server:
   ```
   npm run serve
   ```

   This will start the server on port 3010, accessible at https://175.176.185.170:3010

   Ensure that `utility.sedl.in` points to `175.176.185.170`, and set up a reverse proxy (e.g., nginx) to proxy HTTPS requests from `utility.sedl.in` to `http://localhost:3010`.

   Example nginx config:
   ```
   server {
       listen 443 ssl;
       server_name utility.sedl.in;
       ssl_certificate /path/to/cert.pem;
       ssl_certificate_key /path/to/key.pem;

       location / {
           proxy_pass http://localhost:3010;
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
           proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
           proxy_set_header X-Forwarded-Proto $scheme;
       }
   }
   ```

4. The add-in will be accessible via https://utility.sedl.in

## Development

- `npm run dev-server`: Start development server on port 3010
- `npm run build`: Build for production
- `npm start`: Start debugging in Office