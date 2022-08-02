import { Router, Request, Response } from 'express';
import GenerateReport from './controller';

const router: Router = Router();

router.get('/', async (req: Request, res: Response) => { 
    await GenerateReport.generateReport(req, res);
});

export default router;