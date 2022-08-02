import cluster from 'cluster';
import os from 'os';
import app from './app';
import { resolve } from 'path';
import { config } from 'dotenv';

const PORT: number = Number(process.env.PORT) || 7777;

const noOfCPU = os.cpus().length; // Returns no of cores of cpu

config({ path: resolve(__dirname, '../.env') });

if (cluster.isPrimary) {
    for (let i = 0; i < noOfCPU; i++) {
        cluster.fork();
    }

    cluster.on('exit', (worker, code, signal) => {        
        cluster.fork();
    });
} else {
    app.listen(PORT, () => {
        console.log(`Server is listening on ${PORT}`);
    });
}