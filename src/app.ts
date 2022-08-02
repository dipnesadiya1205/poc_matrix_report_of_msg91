import express, { Application, Request, Response } from 'express';
import router from './routes';

const app: Application = express();

router(app);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.get('/health', (req: Request, res: Response) => {
    console.log('Health Checking');
});

app.all('/*', (req: Request, res: Response) => {
    return res.send('Invalid Routes');
});

export default app;
