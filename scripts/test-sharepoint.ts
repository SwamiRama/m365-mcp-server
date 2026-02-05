#!/usr/bin/env npx tsx
/**
 * Quick test script to verify SharePoint listSites works correctly.
 * Uses device code flow to get a user token interactively.
 *
 * Usage: npx tsx scripts/test-sharepoint.ts
 */

import 'dotenv/config';
import { PublicClientApplication, DeviceCodeRequest } from '@azure/msal-node';

const clientId = process.env['AZURE_CLIENT_ID'];
const tenantId = process.env['AZURE_TENANT_ID'];

if (!clientId || !tenantId) {
  console.error('Missing required environment variables: AZURE_CLIENT_ID, AZURE_TENANT_ID');
  process.exit(1);
}

const msalConfig = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
  },
};

const scopes = [
  'https://graph.microsoft.com/User.Read',
  'https://graph.microsoft.com/Sites.Read.All',
];

async function main() {
  const pca = new PublicClientApplication(msalConfig);

  console.log('\n📱 Starting device code authentication...\n');

  const deviceCodeRequest: DeviceCodeRequest = {
    scopes,
    deviceCodeCallback: (response) => {
      console.log('━'.repeat(60));
      console.log(response.message);
      console.log('━'.repeat(60));
    },
  };

  try {
    const authResult = await pca.acquireTokenByDeviceCode(deviceCodeRequest);

    if (!authResult) {
      console.error('Failed to acquire token');
      process.exit(1);
    }

    console.log('\n✅ Authentication successful!');
    console.log(`   User: ${authResult.account?.username}`);

    // Test the SharePoint sites endpoint with search=*
    console.log('\n🔍 Testing /sites?search=* endpoint...\n');

    const response = await fetch(
      'https://graph.microsoft.com/v1.0/sites?search=*&$top=25&$select=id,name,displayName,webUrl,description',
      {
        headers: {
          Authorization: `Bearer ${authResult.accessToken}`,
        },
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`❌ Graph API error: ${response.status} ${response.statusText}`);
      console.error(errorText);
      process.exit(1);
    }

    const data = await response.json();
    const sites = data.value || [];

    console.log(`✅ Found ${sites.length} SharePoint site(s):\n`);

    if (sites.length === 0) {
      console.log('   (No sites found - user may not have access to any SharePoint sites)');
    } else {
      for (const site of sites) {
        console.log(`   📁 ${site.displayName || site.name}`);
        console.log(`      ID: ${site.id}`);
        console.log(`      URL: ${site.webUrl}`);
        if (site.description) {
          console.log(`      Desc: ${site.description}`);
        }
        console.log('');
      }
    }

    // Also test without search parameter to confirm it returns empty
    console.log('\n🔍 Testing /sites WITHOUT search parameter (should fail or return empty)...\n');

    const response2 = await fetch(
      'https://graph.microsoft.com/v1.0/sites?$top=25&$select=id,name,displayName,webUrl,description',
      {
        headers: {
          Authorization: `Bearer ${authResult.accessToken}`,
        },
      }
    );

    if (!response2.ok) {
      console.log(`   ⚠️  Request failed as expected: ${response2.status} ${response2.statusText}`);
    } else {
      const data2 = await response2.json();
      const sites2 = data2.value || [];
      console.log(`   Result: ${sites2.length} site(s) - ${sites2.length === 0 ? 'Confirmed: empty without search param' : 'Unexpected: got results'}`);
    }

    console.log('\n✅ Test completed successfully!\n');

  } catch (error) {
    console.error('Error:', error);
    process.exit(1);
  }
}

main();
