# Azure Deployment Guide

This guide covers deploying m365-mcp-server to Azure using Container Apps or Web App for Containers.

## Prerequisites

- Azure CLI installed and logged in
- Docker image pushed to a container registry (GitHub Container Registry, Azure Container Registry, or Docker Hub)
- Azure AD app registration completed (see [entra-app-registration.md](./entra-app-registration.md))

## Option 1: Azure Container Apps (Recommended)

Azure Container Apps provides serverless containers with built-in scaling and HTTPS.

### 1. Create Resource Group

```bash
az group create \
  --name m365-mcp-rg \
  --location eastus
```

### 2. Create Container Apps Environment

```bash
az containerapp env create \
  --name m365-mcp-env \
  --resource-group m365-mcp-rg \
  --location eastus
```

### 3. Create Azure Key Vault for Secrets

```bash
# Create Key Vault
az keyvault create \
  --name m365-mcp-kv \
  --resource-group m365-mcp-rg \
  --location eastus

# Add secrets
az keyvault secret set \
  --vault-name m365-mcp-kv \
  --name AZURE-CLIENT-SECRET \
  --value "your-client-secret"

az keyvault secret set \
  --vault-name m365-mcp-kv \
  --name SESSION-SECRET \
  --value "$(openssl rand -hex 32)"
```

### 4. Deploy Container App

```bash
az containerapp create \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --environment m365-mcp-env \
  --image ghcr.io/anthropic/m365-mcp-server:latest \
  --target-port 3000 \
  --ingress external \
  --min-replicas 1 \
  --max-replicas 10 \
  --cpu 0.5 \
  --memory 1Gi \
  --env-vars \
    "AZURE_CLIENT_ID=your-client-id" \
    "AZURE_TENANT_ID=your-tenant-id" \
    "NODE_ENV=production" \
    "LOG_LEVEL=info" \
  --secrets \
    "azure-client-secret=keyvaultref:https://m365-mcp-kv.vault.azure.net/secrets/AZURE-CLIENT-SECRET,identityref:/subscriptions/{sub-id}/resourcegroups/m365-mcp-rg/providers/Microsoft.ManagedIdentity/userAssignedIdentities/m365-mcp-identity" \
    "session-secret=keyvaultref:https://m365-mcp-kv.vault.azure.net/secrets/SESSION-SECRET,identityref:/subscriptions/{sub-id}/resourcegroups/m365-mcp-rg/providers/Microsoft.ManagedIdentity/userAssignedIdentities/m365-mcp-identity"
```

### 5. Get Application URL

```bash
az containerapp show \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --query properties.configuration.ingress.fqdn \
  --output tsv
```

### 6. Update Azure AD Redirect URIs

Add the Container App URL to your Azure AD app registration:
- `https://{your-app-name}.{region}.azurecontainerapps.io/auth/callback`

## Option 2: Azure Web App for Containers

### 1. Create App Service Plan

```bash
az appservice plan create \
  --name m365-mcp-plan \
  --resource-group m365-mcp-rg \
  --is-linux \
  --sku B1
```

### 2. Create Web App

```bash
az webapp create \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --plan m365-mcp-plan \
  --deployment-container-image-name ghcr.io/anthropic/m365-mcp-server:latest
```

### 3. Configure Environment Variables

```bash
az webapp config appsettings set \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --settings \
    AZURE_CLIENT_ID="your-client-id" \
    AZURE_TENANT_ID="your-tenant-id" \
    NODE_ENV="production" \
    WEBSITES_PORT="3000"
```

### 4. Configure Secrets (Use Key Vault References)

```bash
# First, enable managed identity
az webapp identity assign \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg

# Grant Key Vault access
az keyvault set-policy \
  --name m365-mcp-kv \
  --object-id $(az webapp identity show --name m365-mcp-server --resource-group m365-mcp-rg --query principalId -o tsv) \
  --secret-permissions get

# Set secrets as Key Vault references
az webapp config appsettings set \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --settings \
    AZURE_CLIENT_SECRET="@Microsoft.KeyVault(VaultName=m365-mcp-kv;SecretName=AZURE-CLIENT-SECRET)" \
    SESSION_SECRET="@Microsoft.KeyVault(VaultName=m365-mcp-kv;SecretName=SESSION-SECRET)"
```

## Adding Redis for Session Storage

For production with multiple instances, use Azure Cache for Redis:

### 1. Create Redis Cache

```bash
az redis create \
  --name m365-mcp-redis \
  --resource-group m365-mcp-rg \
  --location eastus \
  --sku Basic \
  --vm-size c0
```

### 2. Get Connection String

```bash
az redis show \
  --name m365-mcp-redis \
  --resource-group m365-mcp-rg \
  --query hostName \
  --output tsv
```

### 3. Add to Container App

```bash
az containerapp update \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --set-env-vars "REDIS_URL=redis://m365-mcp-redis.redis.cache.windows.net:6380?ssl=true"
```

## Custom Domain and SSL

### 1. Add Custom Domain

```bash
az containerapp hostname add \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --hostname mcp.yourdomain.com
```

### 2. Configure DNS

Add a CNAME record pointing to your Container App FQDN.

### 3. Enable Managed Certificate

Azure Container Apps automatically provisions SSL certificates for custom domains.

## Monitoring

### Enable Application Insights

```bash
# Create Application Insights
az monitor app-insights component create \
  --app m365-mcp-insights \
  --resource-group m365-mcp-rg \
  --location eastus

# Get instrumentation key
az monitor app-insights component show \
  --app m365-mcp-insights \
  --resource-group m365-mcp-rg \
  --query instrumentationKey \
  --output tsv
```

### Configure Log Analytics

```bash
az containerapp update \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --logs-destination log-analytics \
  --logs-workspace-id <workspace-id>
```

## Scaling Configuration

### Auto-scaling Rules

```bash
az containerapp update \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --min-replicas 1 \
  --max-replicas 10 \
  --scale-rule-name http-scaling \
  --scale-rule-type http \
  --scale-rule-http-concurrency 100
```

## Security Checklist

- [ ] Secrets stored in Azure Key Vault
- [ ] Managed Identity enabled for Key Vault access
- [ ] HTTPS enforced (automatic in Container Apps)
- [ ] Custom domain with SSL certificate
- [ ] Network security groups configured
- [ ] Azure AD redirect URIs updated
- [ ] Application Insights enabled
- [ ] Log Analytics configured
- [ ] Auto-scaling configured
- [ ] Health probes configured

## Troubleshooting

### View Logs

```bash
az containerapp logs show \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --follow
```

### Check Health

```bash
curl https://{your-app-url}/health
```

### Restart App

```bash
az containerapp revision restart \
  --name m365-mcp-server \
  --resource-group m365-mcp-rg \
  --revision <revision-name>
```
