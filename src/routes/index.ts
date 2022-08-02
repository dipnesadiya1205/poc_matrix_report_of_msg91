import { Application } from 'express';
import route from '../component/GenerateReport/v1/route';

export default (app: Application) => {
    app.use('/generate-report', route);
};
