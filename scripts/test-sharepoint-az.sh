#!/bin/bash
# Test SharePoint sites endpoint using Azure CLI for authentication
# Usage: ./scripts/test-sharepoint-az.sh

set -e

echo ""
echo "🔐 Getting access token via Azure CLI..."
TOKEN=$(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)

if [ -z "$TOKEN" ]; then
  echo "❌ Failed to get access token"
  exit 1
fi

echo "✅ Got access token"
echo ""

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🔍 Test 1: /sites?search=* (should return sites)"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

RESULT1=$(curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites?\$search=*&\$top=10&\$select=id,name,displayName,webUrl,description")

echo "$RESULT1" | jq '.'

COUNT1=$(echo "$RESULT1" | jq '.value | length')
echo ""
echo "📊 Result: Found $COUNT1 site(s) with search=*"
echo ""

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🔍 Test 2: /sites WITHOUT search param (should be empty/error)"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

RESULT2=$(curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites?\$top=10&\$select=id,name,displayName,webUrl,description")

echo "$RESULT2" | jq '.'

COUNT2=$(echo "$RESULT2" | jq '.value | length // 0')
echo ""
echo "📊 Result: Found $COUNT2 site(s) without search param"
echo ""

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📋 Summary"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
if [ "$COUNT1" -gt 0 ]; then
  echo "✅ With search=*: $COUNT1 sites found - FIX WORKS!"
else
  echo "⚠️  With search=*: 0 sites found - User may not have SharePoint access"
fi

if [ "$COUNT2" -eq 0 ]; then
  echo "✅ Without search: 0 sites - Confirms the bug in old code"
else
  echo "⚠️  Without search: $COUNT2 sites - Unexpected"
fi
echo ""
