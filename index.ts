import dotenv from 'dotenv';
import webserver from "./web"

dotenv.config();
console.log(webserver.host);