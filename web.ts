// @ts-ignore
import express, { Express, Request, Response , Application } from 'express';
import settings from 'config';
import middleware from 'middleware';

const app: Express = express();

app.get("/", async function(req: Request, res: Response) {
    res.status(500).send("Hello World")
})

app.use("*", async (req, res) => {
    middleware(req, res)
    res.send("test");
})

app.listen(settings.port);
console.log("testing")

export default app;
