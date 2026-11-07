const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const projectRoot = path.resolve(__dirname, '..');
const scriptName = fs.readdirSync(projectRoot)
  .find((name) => name.startsWith('Parametr') && name.endsWith('Ambiente.js'));
const scriptSource = fs.readFileSync(path.join(projectRoot, scriptName), 'utf8');

function createAppsScriptContext(properties, fetchResponse) {
  const calls = [];
  const writes = [];
  const sheet = {
    getRange(range) {
      return {
        setValues(values) {
          writes.push({ range, values: JSON.parse(JSON.stringify(values)) });
        }
      };
    }
  };
  const context = {
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(name) {
            return Object.hasOwn(properties, name) ? properties[name] : null;
          }
        };
      }
    },
    UrlFetchApp: {
      fetch(url, options) {
        calls.push({ url, options });
        return { getContentText: () => JSON.stringify(fetchResponse) };
      }
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName: () => sheet,
          insertSheet: () => sheet
        };
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(scriptSource, context);
  return { context, calls, writes };
}

test('Salesforce configuration missing properties fails before a remote call', () => {
  const { context, calls } = createAppsScriptContext({
    SALESFORCE_CLIENT_ID: 'sf-id'
  }, {});

  assert.throws(() => context.getSalesforceAccessToken(), (error) => {
    assert.match(error.message, /SALESFORCE_CLIENT_SECRET/);
    assert.match(error.message, /SALESFORCE_USERNAME/);
    assert.match(error.message, /SALESFORCE_PASSWORD/);
    assert.match(error.message, /SALESFORCE_SECURITY_TOKEN/);
    assert.match(error.message, /Project Settings > Script Properties/);
    assert.doesNotMatch(error.message, /sf-id/);
    return true;
  });
  assert.equal(calls.length, 0);
});

test('Salesforce uses configured properties in the existing OAuth request', () => {
  const properties = {
    SALESFORCE_CLIENT_ID: 'sf-id',
    SALESFORCE_CLIENT_SECRET: 'sf-secret',
    SALESFORCE_USERNAME: 'sf-user',
    SALESFORCE_PASSWORD: 'sf-pass',
    SALESFORCE_SECURITY_TOKEN: 'sf-token'
  };
  const { context, calls } = createAppsScriptContext(properties, {
    access_token: 'simulated-salesforce-token',
    instance_url: 'https://instance.example.test'
  });

  assert.deepEqual(
    JSON.parse(JSON.stringify(context.getSalesforceAccessToken())),
    { token: 'simulated-salesforce-token', instance: 'https://instance.example.test' }
  );
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, 'https://login.salesforce.com/services/oauth2/token');
  assert.equal(calls[0].options.method, 'post');
  assert.deepEqual(JSON.parse(JSON.stringify(calls[0].options.payload)), {
    grant_type: 'password',
    client_id: properties.SALESFORCE_CLIENT_ID,
    client_secret: properties.SALESFORCE_CLIENT_SECRET,
    username: properties.SALESFORCE_USERNAME,
    password: properties.SALESFORCE_PASSWORD + properties.SALESFORCE_SECURITY_TOKEN
  });
});

test('Lovable configuration missing properties fails before a remote call', () => {
  const { context, calls } = createAppsScriptContext({}, {});

  assert.throws(() => context.realizarLoginEAtualizarApoio(), (error) => {
    assert.match(error.message, /LOVABLE_EMAIL/);
    assert.match(error.message, /LOVABLE_PASSWORD/);
    assert.match(error.message, /LOVABLE_API_KEY_ANON/);
    assert.match(error.message, /Project Settings > Script Properties/);
    return true;
  });
  assert.equal(calls.length, 0);
});

test('Lovable login uses configured properties and preserves the support-sheet token', () => {
  const properties = {
    LOVABLE_EMAIL: 'lov-user',
    LOVABLE_PASSWORD: 'lov-pass',
    LOVABLE_API_KEY_ANON: 'anon-key'
  };
  const { context, calls, writes } = createAppsScriptContext(properties, {
    access_token: 'simulated-lovable-token'
  });

  assert.equal(context.realizarLoginEAtualizarApoio(), 'simulated-lovable-token');
  assert.equal(context.getLovableApiKey(), properties.LOVABLE_API_KEY_ANON);
  assert.equal(calls.length, 1);
  assert.match(calls[0].url, /\/auth\/v1\/token\?grant_type=password$/);
  assert.equal(calls[0].options.method, 'post');
  assert.equal(calls[0].options.contentType, 'application/json');
  assert.equal(calls[0].options.headers.apikey, properties.LOVABLE_API_KEY_ANON);
  assert.deepEqual(JSON.parse(calls[0].options.payload), {
    email: properties.LOVABLE_EMAIL,
    password: properties.LOVABLE_PASSWORD
  });
  assert.deepEqual(writes, [
    { range: 'A1:B1', values: [['CHAVE', 'VALOR']] },
    { range: 'A2:B2', values: [['ACCESS_TOKEN', 'simulated-lovable-token']] }
  ]);
});

test('Lovable API requests use the Script Properties key accessor', () => {
  for (const name of ['GetLovableOrdensService.js', 'OrdensFinalizadas.js', 'Uploads.js']) {
    const source = fs.readFileSync(path.join(projectRoot, name), 'utf8');
    assert.match(source, /getLovableApiKey\(\)/, `${name} should retrieve the configured key`);
    assert.doesNotMatch(source, /\bAPI_KEY_ANON\b/, `${name} should not embed the key`);
  }
});
