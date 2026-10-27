import { reportStartupError, startBridge } from './startup.js';

void startBridge().catch(reportStartupError);
