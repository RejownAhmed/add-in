export default function getBaseUrl(env) {
  switch (env) {
    case 'production':
      return "https://reachapi.reach.app";
    case 'development':
      return "https://apidev.reach.app";
    default:
      return "https://reach-api.test";
  }
}