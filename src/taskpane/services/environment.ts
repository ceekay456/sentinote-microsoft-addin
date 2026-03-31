export interface Environment {
  cognito: {
    userPoolId: string;
    clientId: string;
    domain: string;
    redirectUri: string;
    logoutUri: string;
  };
  apiUrl: string;
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
  },
};

// Default to development; override via webpack DefinePlugin or manual switch
const envName = (typeof process !== "undefined" && process.env?.ADDIN_ENV) || "development";

export const environment: Environment = environments[envName] || environments.development;
