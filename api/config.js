export default function handler(request, response) {
  // Allow cross-origin requests if needed, but normally frontend and API are on the same domain
  response.setHeader('Access-Control-Allow-Origin', '*');
  response.setHeader('Access-Control-Allow-Methods', 'GET');
  
  response.status(200).json({
    apiUrl: process.env.API || ""
  });
}
