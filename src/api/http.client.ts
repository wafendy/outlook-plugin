import axios from "axios";

console.log("BASE_URL", process.env.BASE_URL);

export const api = axios.create({ baseURL: process.env.BASE_URL });
