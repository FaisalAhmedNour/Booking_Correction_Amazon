import { AppCreator } from './utils/App';
import { login } from './functions/login';
import issueBooking from './functions/issueBooking';
import { DataManager } from './utils/DataManager';
import * as path from 'path';
import * as fs from 'fs';

// Load config from EngineTestConfig.json or fallback
let config: any = {};
const configPath = path.join(__dirname, 'EngineTestConfig.json');
if (fs.existsSync(configPath)) {
    try {
        config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    } catch (e) {
        console.warn('⚠️ Could not parse EngineTestConfig.json, using defaults.');
    }
}

async function runLocalTest() {
    console.log('==================================================');
    console.log('🚀 Starting Century VMS Local Booking Correction Test');
    console.log('==================================================');

    const app = new AppCreator();

    // 1. Resolve Excel File Path
    let excelPath = config.initData?.excelFile || path.join(__dirname, 'devDoc', 'CENTURY VMS.xlsx');
    if (!fs.existsSync(excelPath)) {
        excelPath = path.join(__dirname, 'devDoc', 'CENTURY VMS.xlsx');
    }

    if (!fs.existsSync(excelPath)) {
        console.error(`❌ Excel file not found at: ${excelPath}`);
        process.exit(1);
    }
    console.log(`📁 Excel file: ${excelPath}`);

    // 2. Setup Init Data
    app.initData = {
        userName: config.initData?.userName || 'AMZ1599-0',
        password: config.initData?.password || 'Epic@123456',
        excelFile: excelPath,
        settings: {
            companyCode: config.initData?.settings?.companyCode || '589 - Amazon Multi Country Consolidation Thailand',
            shipVendorCode: config.initData?.settings?.shipVendorCode || ' EPLYL - Epic Garments DWC LLC|EPIC Garments Manufacturing Co. Ltd (S004877)|ST-636242|PBGO-Laem Chabang|WW|Private Brands',
            contactPhone: config.initData?.settings?.contactPhone || '01914434666',
            contactEmail: config.initData?.settings?.contactEmail || 'alauddinc@epicbd.com',
            addEditManufacturer: config.initData?.settings?.addEditManufacturer || 'Epic Garments Manufacturing Co. Ltd.',
            emailAddressToReceiveFCR: config.initData?.settings?.emailAddressToReceiveFCR || 'alauddinc@epicbd.com',
            shipperName: config.initData?.settings?.shipperName || 'Epic Garments DWC LLC',
            status: config.initData?.settings?.status || 'Sent',       // Set to 'Sent' to trigger submission
            reviewMode: config.initData?.settings?.reviewMode || false  // Set to false to perform actual submit
        }
    };

    console.log(`👤 User: ${app.initData.userName}`);
    console.log(`📋 Status Mode: ${app.initData.settings.status}`);
    console.log(`🔍 Review Mode: ${app.initData.settings.reviewMode}`);

    // 3. Mock LocalCom to display logs locally
    app.LocalCom = {
        CLIargs: {
            devFlag: true,
            token: 'dev',
            info: '',
            dir: '',
            driver: '',
            localRoot: 'http://127.0.0.1:4116'
        },
        progress: {
            channel: 'd',
            setTotal: (total: number) => {
                console.log(`[PROGRESS] Total items to process: ${total}`);
            },
            getNewSession: (bufferSize: number = 1, applyCharge: boolean = false) => {
                return Promise.resolve({
                    id: 'local_dev_session',
                    remaining: 999999,
                    name: 'OK',
                    message: 'No charge',
                    doNext: true
                });
            },
            setTitleAndMessage: (title: string, details: string) => {
                console.log(`[STATUS] ${title} - ${details || ''}`);
            },
            registerSuccess: (count: number = 1) => {
                console.log(`[SUCCESS] Processed ${count} item(s)`);
                return Promise.resolve({ continue: true, message: 'OK' });
            },
            registerFail: (count: number = 1) => {
                console.log(`[FAIL] Failed ${count} item(s)`);
                return Promise.resolve({ continue: true, message: 'Fail' });
            },
            getObj: () => ({ channel: 'd' })
        },
        sendData: (tableData: any, messages: any, log: any) => {
            if (messages && messages.length > 0) {
                for (const m of messages) {
                    console.log(`[MSG] [${m.type?.toUpperCase() || 'INFO'}] ${m.title}: ${m.message}`);
                }
            }
            if (tableData && tableData.length > 0) {
                // console.log(`[TABLE DATA]`, JSON.stringify(tableData, null, 2));
            }
            return Promise.resolve();
        }
    } as any;

    // 4. Initialize DataManager
    console.log('🔄 Initializing Excel DataManager...');
    app.dateManager = new DataManager(app.initData.excelFile);
    await app.dateManager.init();
    console.log(`✅ Loaded ${app.dateManager.length} row(s) from Excel.`);
    app.verified = true;

    // 5. Launch Browser
    console.log('🌐 Launching Microsoft Edge...');
    await app.launchBrowser();

    // 6. Login
    console.log('🔑 Logging into Century VMS portal...');
    app.LocalCom.progress.setTitleAndMessage('Logging in.', 'Trying to log in to the portal.');
    const loginStatus = await login(app);
    console.log('Login Result:', loginStatus);

    if (!loginStatus.success) {
        console.error('❌ Login failed:', loginStatus.message);
        await app.Browser.close();
        process.exit(1);
    }

    console.log('✅ Login successful! Starting booking generation and submission...');
    app.LocalCom.progress.setTitleAndMessage('Logged-in.', 'Login Success.');

    // 7. Execute Booking & Submission Flow
    await issueBooking(app);

    console.log('🎉 Booking process completed.');
}

runLocalTest().catch((err) => {
    console.error('💥 Uncaught error during local test:', err);
});
