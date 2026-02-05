#!/bin/bash
# Local OAuth test script for MCP Server
# Tests the full OAuth flow including SharePoint access

set -e

CLIENT_ID="a24e4bc6d00dbafab1d9c3b74292fd5e"
REDIRECT_URI="http://localhost:8888/callback"
MCP_SERVER="http://localhost:3000"

# Generate PKCE
CODE_VERIFIER=$(openssl rand -base64 32 | tr -d '=/+' | cut -c1-43)
CODE_CHALLENGE=$(echo -n "$CODE_VERIFIER" | openssl sha256 -binary | base64 | tr '+/' '-_' | tr -d '=')

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🔐 MCP Server Local OAuth Test"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "Step 1: Open this URL in your browser to authenticate:"
echo ""
echo "${MCP_SERVER}/authorize?response_type=code&client_id=${CLIENT_ID}&redirect_uri=${REDIRECT_URI}&scope=openid%20offline_access%20mail.read%20files.read&code_challenge=${CODE_CHALLENGE}&code_challenge_method=S256&state=test123"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "Step 2: After Azure AD login, you'll be redirected to:"
echo "        http://localhost:8888/callback?code=XXX&state=test123"
echo ""
echo "        (The browser will show an error - that's OK!)"
echo "        Copy the 'code' parameter from the URL."
echo ""
read -p "Paste the authorization code here: " AUTH_CODE
echo ""

if [ -z "$AUTH_CODE" ]; then
  echo "❌ No code provided"
  exit 1
fi

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "Step 3: Exchanging code for tokens..."
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

TOKEN_RESPONSE=$(curl -s -X POST "${MCP_SERVER}/token" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "grant_type=authorization_code" \
  -d "code=${AUTH_CODE}" \
  -d "redirect_uri=${REDIRECT_URI}" \
  -d "client_id=${CLIENT_ID}" \
  -d "code_verifier=${CODE_VERIFIER}")

echo "$TOKEN_RESPONSE" | jq '.'

ACCESS_TOKEN=$(echo "$TOKEN_RESPONSE" | jq -r '.access_token // empty')

if [ -z "$ACCESS_TOKEN" ]; then
  echo ""
  echo "❌ Failed to get access token"
  exit 1
fi

echo ""
echo "✅ Got access token!"
echo ""

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "Step 4: Testing MCP tools/list endpoint..."
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

curl -s -X POST "${MCP_SERVER}/mcp" \
  -H "Authorization: Bearer ${ACCESS_TOKEN}" \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","method":"tools/list","id":1}' | jq '.result.tools[].name'

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "Step 5: Testing sp_list_sites (SharePoint)..."
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

RESULT=$(curl -s -X POST "${MCP_SERVER}/mcp" \
  -H "Authorization: Bearer ${ACCESS_TOKEN}" \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","method":"tools/call","params":{"name":"sp_list_sites","arguments":{}},"id":2}')

echo "$RESULT" | jq '.'

# Extract count
SITE_COUNT=$(echo "$RESULT" | jq -r '.result.content[0].text' 2>/dev/null | jq -r '.count // 0' 2>/dev/null || echo "0")

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📋 Result"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
if [ "$SITE_COUNT" -gt 0 ] 2>/dev/null; then
  echo "✅ SUCCESS! Found ${SITE_COUNT} SharePoint site(s)"
else
  echo "⚠️  Found 0 sites (user may not have SharePoint access, or consent is needed)"
fi
echo ""
