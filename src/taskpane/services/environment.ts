export interface Environment {
  cognito: {
    userPoolId: string;
    clientId: string;
    domain: string;
    redirectUri: string;
    logoutUri: string;
  };
  apiUrl: string;
  microsoft: {
    clientId: string;
    tenantId: string;
    redirectUri: string;
    scope: string;
  };
}

const environments: Record<string, Environment> = {
  development: {
    cognito: {
      userPoolId: "eu-central-1_pko7ADLrF",
      clientId: "vi1t93csn957hu7nvbuh6i3bb",
      domain: "sentinote-612559614330.auth.eu-central-1.amazoncognito.com",
      redirectUri: "https://localhost:3000/callback.html",
      logoutUri: "https://localhost:3000",
    },
    apiUrl: "https://86rcreh9mh.execute-api.eu-central-1.amazonaws.com/prod",
    microsoft: {
      clientId: "5bc3d59a-71d4-4780-9f0e-31520a6797a6",
      tenantId: "fedcef2f-0c85-40dd-8f55-e23143dcb367",
      redirectUri: "https://localhost:3000/ms-callback.html",
      scope: "https://graph.microsoft.com/Files.ReadWrite",
    },
  },
  freshminds: {
    cognito: {
      userPoolId: "eu-central-1_L064ymPKl",
      clientId: "7ccmiibl2imaooo7gafm60ai7c",
      domain: "researchcloud-931097097198.auth.eu-central-1.amazoncognito.com",
      redirectUri: "https://localhost:3000/callback.html",
      logoutUri: "https://localhost:3000",
    },
    apiUrl: "https://vydindm26e.execute-api.eu-central-1.amazonaws.com/prod",
    microsoft: {
      clientId: "5bc3d59a-71d4-4780-9f0e-31520a6797a6",
      tenantId: "fedcef2f-0c85-40dd-8f55-e23143dcb367",
      redirectUri: "https://researchcloud.app/msoffice-addin/ms-callback.html",
      scope: "https://graph.microsoft.com/Files.ReadWrite",
    },
  },
};

// Replaced at build time by webpack.DefinePlugin (see webpack.config.js)
const envName = process.env.ADDIN_ENV || "development";

export const environment: Environment = environments[envName] || environments.development;
