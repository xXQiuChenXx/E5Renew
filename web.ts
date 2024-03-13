// @ts-ignore
import express, { Express, Request, Response , Application } from 'web';
import settings from 'config';

const app: Express = express();

app.get("/", async function(req: Request, res: Response) {
    res.status(500).send("Hello World")
})

app.listen(settings.port);
console.log("testing")

export default app;
