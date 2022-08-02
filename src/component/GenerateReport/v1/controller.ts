import { Request, Response } from 'express';
import ExcelJS from 'exceljs';
import STATUS_CODE from 'http-status-codes';
import moment from 'moment';
import * as xlsx from 'xlsx';
import { createResponse } from '../../../utils/helper';

class GenerateReport {
    public certificateIssueFile;
    public forgotPasswordFile;
    public resendCertificateFile;
    public updateContactNoFile;
    public verificationOTP2FAFile;
    public skillPassVerificationOTPFile;
    public skillPassForgotPasswordFile;
    public skillPassUpdateContactNoFile;

    constructor() {
        this.certificateIssueFile = xlsx.readFile(
            '/home/dipkumar/Matrix Report using Excel/src/report/certificate_issue.xlsx'
        );
        this.forgotPasswordFile = xlsx.readFile(
            '/home/dipkumar/Matrix Report using Excel/src/report/forgot_password.xlsx'
        );
        this.resendCertificateFile = xlsx.readFile(
            '/home/dipkumar/Matrix Report using Excel/src/report/resend_certificate.xlsx'
        );
        this.updateContactNoFile = xlsx.readFile(
            '/home/dipkumar/Matrix Report using Excel/src/report/update_contact_no.xlsx'
        );
        this.verificationOTP2FAFile = xlsx.readFile(
            '/home/dipkumar/Matrix Report using Excel/src/report/verification_otp_2FA.xlsx'
        );
        this.skillPassVerificationOTPFile = xlsx.readFile(
            '/home/dipkumar/Matrix Report using Excel/src/report/skillpass_verification_otp.xlsx'
        );
        this.skillPassForgotPasswordFile = xlsx.readFile(
            '/home/dipkumar/Matrix Report using Excel/src/report/skillpass_forgot_password.xlsx'
        );
        this.skillPassUpdateContactNoFile = xlsx.readFile(
            '/home/dipkumar/Matrix Report using Excel/src/report/skillpass_update_contact_no.xlsx'
        );
    }

    public async getCertificateIssueData(req: Request, res: Response) {
        try {
            const { sheet_name: sheetName, start_date: startDate } = req.query;

            let worksheet = this.certificateIssueFile.Sheets[`${sheetName}`];

            let data: any = xlsx.utils.sheet_to_json(worksheet);

            let weekdates: string[] = [];

            for (let i = 0; i <= 6; i++) {
                let today = moment(`${startDate}`).format('YYYY-MM-DD');
                weekdates.push(moment(today).add(i, 'days').format('YYYY-MM-DD'));
            }

            let newData: any[] = [];

            if (data.length === 0) {
                weekdates.forEach((el: any) => {
                    let obj: any = {};
                    obj[el] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                });
                return newData;
            }
            /*
            let newData = [
                {
                  '2022-04-14': { Delivered: 0, Rejected: 4, Failed: 0, Total: 4, Credit: 0 }
                },
                {
                  '2022-04-15': { Delivered: 0, Rejected: 2, Failed: 0, Total: 2, Credit: 0 }
                },
            ]    
            */

            let temp: string[] = []; // Prevent from making duplicate key in newData

            data.forEach((el: any) => {
                let date: string = el.SentTime.split(' ')[0];

                if (weekdates.includes(date)) {
                    if (!temp.includes(date)) {
                        let obj: any = {};

                        if (el.Status === 'Delivered') {
                            obj[date] = {
                                Delivered: 1,
                                Rejected: 0, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else if (el.Status === 'Failed' || el.Status === 'Report Pending') {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 0, // Blocked Number
                                Failed: 1, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 1, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        }

                        newData.push(obj);
                        temp.push(date);
                    } else {
                        newData.forEach((element) => {
                            if (Object.keys(element)[0] === date) {
                                if (el.Status === 'Delivered') {
                                    element[date].Delivered = Number(element[date].Delivered) + 1;
                                } else if (
                                    el.Status === 'Report Pending' ||
                                    el.Status === 'Failed'
                                ) {
                                    element[date].Failed = Number(element[date].Failed) + 1;
                                } else {
                                    element[date].Rejected = Number(element[date].Rejected) + 1;
                                }
                                element[date].Total = Number(element[date].Total) + 1;
                                element[date].Credit =
                                    Number(element[date].Credit) + Number(el.Credit);
                            }
                        });
                    }
                }
            });
            temp.length = 0;

            let includedDates: string[] = [];

            newData.forEach((el: any) => {
                includedDates.push(Object.keys(el)[0]);
            });

            for (let i = 0; i < weekdates.length; i++) {
                if (!includedDates.includes(weekdates[i])) {
                    let obj: any = {};
                    obj[weekdates[i]] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                }
            }
            weekdates.length = 0;
            includedDates.length = 0;

            return newData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async getForgotPasswordData(req: Request, res: Response) {
        try {
            const { sheet_name: sheetName, start_date: startDate } = req.query;

            let worksheet = this.forgotPasswordFile.Sheets[`${sheetName}`];

            let data: any = xlsx.utils.sheet_to_json(worksheet);

            let weekdates: string[] = [];

            for (let i = 0; i <= 6; i++) {
                let today = moment(`${startDate}`).format('YYYY-MM-DD');
                weekdates.push(moment(today).add(i, 'days').format('YYYY-MM-DD'));
            }

            let newData: any[] = [];

            if (data.length === 0) {
                weekdates.forEach((el: any) => {
                    let obj: any = {};
                    obj[el] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                });
                return newData;
            }

            /*
            let newData = [{
                    '07-04-2022':
                        { Delivered:2, Rejected:2, Total:4 }
                }]
            */

            let temp: string[] = []; // Prevent from making duplicate key in newData

            data.forEach((el: any) => {
                let date: string = el.SentTime.split(' ')[0];

                if (weekdates.includes(date)) {
                    if (!temp.includes(date)) {
                        let obj: any = {};

                        if (el.Status === 'Delivered') {
                            obj[date] = {
                                Delivered: 1,
                                Rejected: 0, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else if (el.Status === 'Failed' || el.Status === 'Report Pending') {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 0, // Blocked Number
                                Failed: 1, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 1, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        }

                        newData.push(obj);
                        temp.push(date);
                    } else {
                        newData.forEach((element) => {
                            if (Object.keys(element)[0] === date) {
                                if (el.Status === 'Delivered') {
                                    element[date].Delivered = Number(element[date].Delivered) + 1;
                                } else if (
                                    el.Status === 'Report Pending' ||
                                    el.Status === 'Failed'
                                ) {
                                    element[date].Failed = Number(element[date].Failed) + 1;
                                } else {
                                    element[date].Rejected = Number(element[date].Rejected) + 1;
                                }
                                element[date].Total = Number(element[date].Total) + 1;
                                element[date].Credit =
                                    Number(element[date].Credit) + Number(el.Credit);
                            }
                        });
                    }
                }
            });
            temp.length = 0;

            let includedDates: string[] = [];

            newData.forEach((el: any) => {
                includedDates.push(Object.keys(el)[0]);
            });

            for (let i = 0; i < weekdates.length; i++) {
                if (!includedDates.includes(weekdates[i])) {
                    let obj: any = {};
                    obj[weekdates[i]] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                }
            }
            weekdates.length = 0;
            includedDates.length = 0;

            return newData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async getResendCertificateData(req: Request, res: Response) {
        try {
            const { sheet_name: sheetName, start_date: startDate } = req.query;

            let worksheet = this.resendCertificateFile.Sheets[`${sheetName}`];

            let data: any = xlsx.utils.sheet_to_json(worksheet);

            let weekdates: string[] = [];

            for (let i = 0; i <= 6; i++) {
                let today = moment(`${startDate}`).format('YYYY-MM-DD');
                weekdates.push(moment(today).add(i, 'days').format('YYYY-MM-DD'));
            }

            let newData: any[] = [];

            if (data.length === 0) {
                weekdates.forEach((el: any) => {
                    let obj: any = {};
                    obj[el] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                });
                return newData;
            }

            /*
            let newData = [{
                    '07-04-2022':
                        { Delivered:2, Rejected:2, Total:4 }
                }]
            */

            let temp: string[] = []; // Prevent from making duplicate key in newData

            data.forEach((el: any) => {
                let date: string = el.SentTime.split(' ')[0];

                if (weekdates.includes(date)) {
                    if (!temp.includes(date)) {
                        let obj: any = {};

                        if (el.Status === 'Delivered') {
                            obj[date] = {
                                Delivered: 1,
                                Rejected: 0, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else if (el.Status === 'Failed' || el.Status === 'Report Pending') {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 0, // Blocked Number
                                Failed: 1, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 1, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        }

                        newData.push(obj);
                        temp.push(date);
                    } else {
                        newData.forEach((element) => {
                            if (Object.keys(element)[0] === date) {
                                if (el.Status === 'Delivered') {
                                    element[date].Delivered = Number(element[date].Delivered) + 1;
                                } else if (
                                    el.Status === 'Report Pending' ||
                                    el.Status === 'Failed'
                                ) {
                                    element[date].Failed = Number(element[date].Failed) + 1;
                                } else {
                                    element[date].Rejected = Number(element[date].Rejected) + 1;
                                }
                                element[date].Total = Number(element[date].Total) + 1;
                                element[date].Credit =
                                    Number(element[date].Credit) + Number(el.Credit);
                            }
                        });
                    }
                }
            });
            temp.length = 0;

            let includedDates: string[] = [];

            newData.forEach((el: any) => {
                includedDates.push(Object.keys(el)[0]);
            });

            for (let i = 0; i < weekdates.length; i++) {
                if (!includedDates.includes(weekdates[i])) {
                    let obj: any = {};
                    obj[weekdates[i]] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                }
            }
            weekdates.length = 0;
            includedDates.length = 0;

            return newData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async getupdateContactNoData(req: Request, res: Response) {
        try {
            const { sheet_name: sheetName, start_date: startDate } = req.query;

            let worksheet = this.updateContactNoFile.Sheets[`${sheetName}`];

            let data: any = xlsx.utils.sheet_to_json(worksheet);

            let weekdates: string[] = [];

            for (let i = 0; i <= 6; i++) {
                let today = moment(`${startDate}`).format('YYYY-MM-DD');
                weekdates.push(moment(today).add(i, 'days').format('YYYY-MM-DD'));
            }

            let newData: any[] = [];

            if (data.length === 0) {
                weekdates.forEach((el: any) => {
                    let obj: any = {};
                    obj[el] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                });
                return newData;
            }

            /*
            let newData = [{
                    '07-04-2022':
                        { Delivered:2, Rejected:2, Total:4 }
                }]
            */

            let temp: string[] = []; // Prevent from making duplicate key in newData

            data.forEach((el: any) => {
                let date: string = el.SentTime.split(' ')[0];

                if (weekdates.includes(date)) {
                    if (!temp.includes(date)) {
                        let obj: any = {};

                        if (el.Status === 'Delivered') {
                            obj[date] = {
                                Delivered: 1,
                                Rejected: 0, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else if (el.Status === 'Failed' || el.Status === 'Report Pending') {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 0, // Blocked Number
                                Failed: 1, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 1, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        }

                        newData.push(obj);
                        temp.push(date);
                    } else {
                        newData.forEach((element) => {
                            if (Object.keys(element)[0] === date) {
                                if (el.Status === 'Delivered') {
                                    element[date].Delivered = Number(element[date].Delivered) + 1;
                                } else if (
                                    el.Status === 'Report Pending' ||
                                    el.Status === 'Failed'
                                ) {
                                    element[date].Failed = Number(element[date].Failed) + 1;
                                } else {
                                    element[date].Rejected = Number(element[date].Rejected) + 1;
                                }
                                element[date].Total = Number(element[date].Total) + 1;
                                element[date].Credit =
                                    Number(element[date].Credit) + Number(el.Credit);
                            }
                        });
                    }
                }
            });
            temp.length = 0;

            let includedDates: string[] = [];

            newData.forEach((el: any) => {
                includedDates.push(Object.keys(el)[0]);
            });

            for (let i = 0; i < weekdates.length; i++) {
                if (!includedDates.includes(weekdates[i])) {
                    let obj: any = {};
                    obj[weekdates[i]] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                }
            }
            weekdates.length = 0;
            includedDates.length = 0;

            return newData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async getVerificationOTP2FAData(req: Request, res: Response) {
        try {
            const { sheet_name: sheetName, start_date: startDate } = req.query;

            let worksheet = this.verificationOTP2FAFile.Sheets[`${sheetName}`];

            let data: any = xlsx.utils.sheet_to_json(worksheet);

            let weekdates: string[] = [];

            for (let i = 0; i <= 6; i++) {
                let today = moment(`${startDate}`).format('YYYY-MM-DD');
                weekdates.push(moment(today).add(i, 'days').format('YYYY-MM-DD'));
            }

            let newData: any[] = [];

            if (data.length === 0) {
                weekdates.forEach((el: any) => {
                    let obj: any = {};
                    obj[el] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                });
                return newData;
            }

            /*
            let newData = [{
                    '07-04-2022':
                        { Delivered:2, Rejected:2, Total:4 }
                }]
            */

            let temp: string[] = []; // Prevent from making duplicate key in newData

            data.forEach((el: any) => {
                let date: string = el.SentTime.split(' ')[0];

                if (weekdates.includes(date)) {
                    if (!temp.includes(date)) {
                        let obj: any = {};

                        if (el.Status === 'Delivered') {
                            obj[date] = {
                                Delivered: 1,
                                Rejected: 0, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else if (el.Status === 'Failed' || el.Status === 'Report Pending') {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 0, // Blocked Number
                                Failed: 1, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 1, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        }

                        newData.push(obj);
                        temp.push(date);
                    } else {
                        newData.forEach((element) => {
                            if (Object.keys(element)[0] === date) {
                                if (el.Status === 'Delivered') {
                                    element[date].Delivered = Number(element[date].Delivered) + 1;
                                } else if (
                                    el.Status === 'Report Pending' ||
                                    el.Status === 'Failed'
                                ) {
                                    element[date].Failed = Number(element[date].Failed) + 1;
                                } else {
                                    element[date].Rejected = Number(element[date].Rejected) + 1;
                                }
                                element[date].Total = Number(element[date].Total) + 1;
                                element[date].Credit =
                                    Number(element[date].Credit) + Number(el.Credit);
                            }
                        });
                    }
                }
            });
            temp.length = 0;

            let includedDates: string[] = [];

            newData.forEach((el: any) => {
                includedDates.push(Object.keys(el)[0]);
            });

            for (let i = 0; i < weekdates.length; i++) {
                if (!includedDates.includes(weekdates[i])) {
                    let obj: any = {};
                    obj[weekdates[i]] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                }
            }
            weekdates.length = 0;
            includedDates.length = 0;

            return newData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async getSkillPassVerificationOTPData(req: Request, res: Response) {
        try {
            const { sheet_name: sheetName, start_date: startDate } = req.query;

            let worksheet = this.skillPassVerificationOTPFile.Sheets[`${sheetName}`];

            let data: any = xlsx.utils.sheet_to_json(worksheet);

            let weekdates: string[] = [];

            for (let i = 0; i <= 6; i++) {
                let today = moment(`${startDate}`).format('YYYY-MM-DD');
                weekdates.push(moment(today).add(i, 'days').format('YYYY-MM-DD'));
            }

            let newData: any[] = [];

            if (data.length === 0) {
                weekdates.forEach((el: any) => {
                    let obj: any = {};
                    obj[el] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                });
                return newData;
            }

            /*
            let newData = [{
                    '07-04-2022':
                        { Delivered:2, Rejected:2, Total:4 }
                }]
            */

            let temp: string[] = []; // Prevent from making duplicate key in newData

            data.forEach((el: any) => {
                let date: string = el.SentTime.split(' ')[0];

                if (weekdates.includes(date)) {
                    if (!temp.includes(date)) {
                        let obj: any = {};

                        if (el.Status === 'Delivered') {
                            obj[date] = {
                                Delivered: 1,
                                Rejected: 0, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else if (el.Status === 'Failed' || el.Status === 'Report Pending') {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 0, // Blocked Number
                                Failed: 1, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 1, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        }

                        newData.push(obj);
                        temp.push(date);
                    } else {
                        newData.forEach((element) => {
                            if (Object.keys(element)[0] === date) {
                                if (el.Status === 'Delivered') {
                                    element[date].Delivered = Number(element[date].Delivered) + 1;
                                } else if (
                                    el.Status === 'Report Pending' ||
                                    el.Status === 'Failed'
                                ) {
                                    element[date].Failed = Number(element[date].Failed) + 1;
                                } else {
                                    element[date].Rejected = Number(element[date].Rejected) + 1;
                                }
                                element[date].Total = Number(element[date].Total) + 1;
                                element[date].Credit =
                                    Number(element[date].Credit) + Number(el.Credit);
                            }
                        });
                    }
                }
            });
            temp.length = 0;

            let includedDates: string[] = [];

            newData.forEach((el: any) => {
                includedDates.push(Object.keys(el)[0]);
            });

            for (let i = 0; i < weekdates.length; i++) {
                if (!includedDates.includes(weekdates[i])) {
                    let obj: any = {};
                    obj[weekdates[i]] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                }
            }
            weekdates.length = 0;
            includedDates.length = 0;

            return newData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async getSkillPassForgotPasswordData(req: Request, res: Response) {
        try {
            const { sheet_name: sheetName, start_date: startDate } = req.query;

            let worksheet = this.skillPassForgotPasswordFile.Sheets[`${sheetName}`];

            let data: any = xlsx.utils.sheet_to_json(worksheet);

            let weekdates: string[] = [];

            for (let i = 0; i <= 6; i++) {
                let today = moment(`${startDate}`).format('YYYY-MM-DD');
                weekdates.push(moment(today).add(i, 'days').format('YYYY-MM-DD'));
            }

            let newData: any[] = [];

            if (data.length === 0) {
                weekdates.forEach((el: any) => {
                    let obj: any = {};
                    obj[el] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                });
                return newData;
            }

            /*
            let newData = [{
                    '07-04-2022':
                        { Delivered:2, Rejected:2, Total:4 }
                }]
            */

            let temp: string[] = []; // Prevent from making duplicate key in newData

            data.forEach((el: any) => {
                let date: string = el.SentTime.split(' ')[0];

                if (weekdates.includes(date)) {
                    if (!temp.includes(date)) {
                        let obj: any = {};

                        if (el.Status === 'Delivered') {
                            obj[date] = {
                                Delivered: 1,
                                Rejected: 0, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else if (el.Status === 'Failed' || el.Status === 'Report Pending') {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 0, // Blocked Number
                                Failed: 1, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 1, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        }

                        newData.push(obj);
                        temp.push(date);
                    } else {
                        newData.forEach((element) => {
                            if (Object.keys(element)[0] === date) {
                                if (el.Status === 'Delivered') {
                                    element[date].Delivered = Number(element[date].Delivered) + 1;
                                } else if (
                                    el.Status === 'Report Pending' ||
                                    el.Status === 'Failed'
                                ) {
                                    element[date].Failed = Number(element[date].Failed) + 1;
                                } else {
                                    element[date].Rejected = Number(element[date].Rejected) + 1;
                                }
                                element[date].Total = Number(element[date].Total) + 1;
                                element[date].Credit =
                                    Number(element[date].Credit) + Number(el.Credit);
                            }
                        });
                    }
                }
            });
            temp.length = 0;

            let includedDates: string[] = [];

            newData.forEach((el: any) => {
                includedDates.push(Object.keys(el)[0]);
            });

            for (let i = 0; i < weekdates.length; i++) {
                if (!includedDates.includes(weekdates[i])) {
                    let obj: any = {};
                    obj[weekdates[i]] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                }
            }
            weekdates.length = 0;
            includedDates.length = 0;

            return newData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async getSkillPassUpdateContactNoData(req: Request, res: Response) {
        try {
            const { sheet_name: sheetName, start_date: startDate } = req.query;

            let worksheet = this.skillPassUpdateContactNoFile.Sheets[`${sheetName}`];

            let data: any = xlsx.utils.sheet_to_json(worksheet);

            let weekdates: string[] = [];

            for (let i = 0; i <= 6; i++) {
                let today = moment(`${startDate}`).format('YYYY-MM-DD');
                weekdates.push(moment(today).add(i, 'days').format('YYYY-MM-DD'));
            }

            let newData: any[] = [];

            if (data.length === 0) {
                weekdates.forEach((el: any) => {
                    let obj: any = {};
                    obj[el] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                });
                return newData;
            }

            /*
            let newData = [{
                    '07-04-2022':
                        { Delivered:2, Rejected:2, Total:4 }
                }]
            */

            let temp: string[] = []; // Prevent from making duplicate key in newData

            data.forEach((el: any) => {
                let date: string = el.SentTime.split(' ')[0];

                if (weekdates.includes(date)) {
                    if (!temp.includes(date)) {
                        let obj: any = {};

                        if (el.Status === 'Delivered') {
                            obj[date] = {
                                Delivered: 1,
                                Rejected: 0, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else if (el.Status === 'Failed' || el.Status === 'Report Pending') {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 0, // Blocked Number
                                Failed: 1, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        } else {
                            obj[date] = {
                                Delivered: 0,
                                Rejected: 1, // Blocked Number
                                Failed: 0, // Report Pending
                                Total: 1,
                                Credit: el.Credit,
                            };
                        }

                        newData.push(obj);
                        temp.push(date);
                    } else {
                        newData.forEach((element) => {
                            if (Object.keys(element)[0] === date) {
                                if (el.Status === 'Delivered') {
                                    element[date].Delivered = Number(element[date].Delivered) + 1;
                                } else if (
                                    el.Status === 'Report Pending' ||
                                    el.Status === 'Failed'
                                ) {
                                    element[date].Failed = Number(element[date].Failed) + 1;
                                } else {
                                    element[date].Rejected = Number(element[date].Rejected) + 1;
                                }
                                element[date].Total = Number(element[date].Total) + 1;
                                element[date].Credit =
                                    Number(element[date].Credit) + Number(el.Credit);
                            }
                        });
                    }
                }
            });
            temp.length = 0;

            let includedDates: string[] = [];

            newData.forEach((el: any) => {
                includedDates.push(Object.keys(el)[0]);
            });

            for (let i = 0; i < weekdates.length; i++) {
                if (!includedDates.includes(weekdates[i])) {
                    let obj: any = {};
                    obj[weekdates[i]] = {
                        Delivered: 0,
                        Rejected: 0, // Blocked Number
                        Failed: 0, // Report Pending
                        Total: 0,
                        Credit: 0,
                    };
                    newData.push(obj);
                }
            }
            weekdates.length = 0;
            includedDates.length = 0;

            return newData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async getTotalSMSData(req: Request, res: Response) {
        try {
            const certificateIssueData: any = await this.getCertificateIssueData(req, res);
            const forgotPasswordData: any = await this.getForgotPasswordData(req, res);
            const resendCertificateData: any = await this.getResendCertificateData(req, res);
            const updateContactNoData: any = await this.getupdateContactNoData(req, res);
            const verificationOTP2FAData: any = await this.getVerificationOTP2FAData(req, res);
            const skillPassVerificationOTPData: any = await this.getSkillPassVerificationOTPData(
                req,
                res
            );
            const skillPassForgotPasswordData: any = await this.getSkillPassForgotPasswordData(
                req,
                res
            );
            const skillPassUpdateContactNoData: any = await this.getSkillPassUpdateContactNoData(
                req,
                res
            );

            let totalSMSData: any[] = [];

            for (let i = 0; i < certificateIssueData.length; i++) {
                let date = Object.keys(certificateIssueData[i])[0];

                let temp1: any = certificateIssueData[i],
                    temp2: any,
                    temp3: any,
                    temp4: any,
                    temp5: any,
                    temp6: any,
                    temp7: any,
                    temp8: any;

                forgotPasswordData.forEach((el: any) => {
                    if (Object.keys(el)[0] === date) {
                        temp2 = el;
                        return;
                    }
                });
                resendCertificateData.forEach((el: any) => {
                    if (Object.keys(el)[0] === date) {
                        temp3 = el;
                        return;
                    }
                });
                updateContactNoData.forEach((el: any) => {
                    if (Object.keys(el)[0] === date) {
                        temp4 = el;
                        return;
                    }
                });
                verificationOTP2FAData.forEach((el: any) => {
                    if (Object.keys(el)[0] === date) {
                        temp5 = el;
                        return;
                    }
                });
                skillPassVerificationOTPData.forEach((el: any) => {
                    if (Object.keys(el)[0] === date) {
                        temp6 = el;
                        return;
                    }
                });
                skillPassForgotPasswordData.forEach((el: any) => {
                    if (Object.keys(el)[0] === date) {
                        temp7 = el;
                        return;
                    }
                });
                skillPassUpdateContactNoData.forEach((el: any) => {
                    if (Object.keys(el)[0] === date) {
                        temp8 = el;
                        return;
                    }
                });
                let totalDelivered =
                    Number(temp1[date].Delivered) +
                    Number(temp2[date].Delivered) +
                    Number(temp3[date].Delivered) +
                    Number(temp4[date].Delivered) +
                    Number(temp5[date].Delivered) +
                    Number(temp6[date].Delivered) +
                    Number(temp7[date].Delivered) +
                    Number(temp8[date].Delivered);
                let totalRejected =
                    Number(temp1[date].Rejected) +
                    Number(temp2[date].Rejected) +
                    Number(temp3[date].Rejected) +
                    Number(temp4[date].Rejected) +
                    Number(temp5[date].Rejected) +
                    Number(temp6[date].Rejected) +
                    Number(temp7[date].Rejected) +
                    Number(temp8[date].Rejected);
                let totalFailed =
                    Number(temp1[date].Failed) +
                    Number(temp2[date].Failed) +
                    Number(temp3[date].Failed) +
                    Number(temp4[date].Failed) +
                    Number(temp5[date].Failed) +
                    Number(temp6[date].Failed) +
                    Number(temp7[date].Failed) +
                    Number(temp8[date].Failed);
                let totalMessages =
                    Number(temp1[date].Total) +
                    Number(temp2[date].Total) +
                    Number(temp3[date].Total) +
                    Number(temp4[date].Total) +
                    Number(temp5[date].Total) +
                    Number(temp6[date].Total) +
                    Number(temp7[date].Total) +
                    Number(temp8[date].Total);
                let totalCredits =
                    Number(temp1[date].Credit) +
                    Number(temp2[date].Credit) +
                    Number(temp3[date].Credit) +
                    Number(temp4[date].Credit) +
                    Number(temp5[date].Credit) +
                    Number(temp6[date].Credit) +
                    Number(temp7[date].Credit) +
                    Number(temp8[date].Credit);
                let obj: any = {};
                obj[date] = {
                    Delivered: totalDelivered,
                    Rejected: totalRejected, // Blocked Number
                    Failed: totalFailed, // Report Pending
                    Total: totalMessages,
                    Credit: totalCredits,
                };
                totalSMSData.push(obj);
            }

            return totalSMSData;
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }

    public async generateReport(req: Request, res: Response) {
        try {
            const workbook = new ExcelJS.Workbook();
            workbook.addWorksheet('MSG91', { properties: { defaultColWidth: 15 } });
            const msg91Worksheet = workbook.getWorksheet('MSG91');
            let row = 3;

            // * ---------------------------  Total SMS  -------------------------------

            const totalSMSData: any = await this.getTotalSMSData(req, res);

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = 'Total SMS';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Certificate Issue

            msg91Worksheet.addTable({
                name: 'table1',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table1 = msg91Worksheet.getTable('table1');

            totalSMSData.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table1.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            table1.commit();
            row++;
            row++;
            row++;

            // * ---------------------------  Certificate Issue  -------------------------------

            let certificateIssueData: any = await this.getCertificateIssueData(req, res);

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = 'Certificate Issue';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Certificate Issue

            msg91Worksheet.addTable({
                name: 'table2',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table2 = msg91Worksheet.getTable('table2');

            certificateIssueData.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table2.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            table2.commit();
            row++;
            row++;
            row++;

            // * ---------------------------  Forgot Password  -------------------------------

            let forgotPasswordData: any = await this.getForgotPasswordData(req, res);

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = 'Forgot Password';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Forgot Password

            msg91Worksheet.addTable({
                name: 'table3',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                    showLastColumn: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table3 = msg91Worksheet.getTable('table3');

            forgotPasswordData.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table3.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;
            table3.commit();
            row++;
            row++;
            row++;

            // * ---------------------------  Resend Certificate  -------------------------------

            let resendCertificateData: any = await this.getResendCertificateData(req, res);

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = 'Resend Certificate';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Forgot Password

            msg91Worksheet.addTable({
                name: 'table4',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table4 = msg91Worksheet.getTable('table4');

            resendCertificateData.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table4.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            table4.commit();
            row++;
            row++;
            row++;

            // * ---------------------------  Update Contact Number  -------------------------------

            let updateContactNoData: any = await this.getupdateContactNoData(req, res);

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = 'Update Contact Number';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Forgot Password

            msg91Worksheet.addTable({
                name: 'table5',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table5 = msg91Worksheet.getTable('table5');

            updateContactNoData.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table5.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;
            table5.commit();
            row++;
            row++;
            row++;

            // * ---------------------------  2FA Verification OTP  -------------------------------

            let OTP2FAVerification: any = await this.getVerificationOTP2FAData(req, res);

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = '2FA Verification OTP';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Forgot Password

            msg91Worksheet.addTable({
                name: 'table6',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table6 = msg91Worksheet.getTable('table6');

            OTP2FAVerification.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table6.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;
            table6.commit();
            row++;
            row++;
            row++;

            // * ---------------------------  Skillpass - 2FA Verification OTP  -------------------------------

            let skillPass2FAData: any = await this.getSkillPassVerificationOTPData(req, res);

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = 'Skillpass - 2FA Verification OTP';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Forgot Password

            msg91Worksheet.addTable({
                name: 'table7',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table7 = msg91Worksheet.getTable('table7');

            skillPass2FAData.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table7.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            table7.commit();
            row++;
            row++;
            row++;

            // * ---------------------------  Skillpass - Forgot Password  -------------------------------

            let skillPassForgotPasswordData: any = await this.getSkillPassForgotPasswordData(
                req,
                res
            );

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = 'Skillpass - Forgot Password';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Forgot Password

            msg91Worksheet.addTable({
                name: 'table8',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table8 = msg91Worksheet.getTable('table8');

            skillPassForgotPasswordData.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table8.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            table8.commit();
            row++;
            row++;
            row++;

            // * ---------------------------  Skillpass - Update Contact Number  -------------------------------

            let skillPassUpdateContactNoData: any = await this.getSkillPassUpdateContactNoData(
                req,
                res
            );

            msg91Worksheet.mergeCells('C' + row + ':H' + row);
            msg91Worksheet.getCell('C' + row).value = 'Skillpass - Update Contact Number';
            msg91Worksheet.getCell('C' + row).alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            msg91Worksheet.getRow(row).font = { bold: true, name: 'calibri' };
            msg91Worksheet.getCell('C' + row).border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
                left: { style: 'thin' },
            };
            row++;

            // Worksheet Table Creation of Forgot Password

            msg91Worksheet.addTable({
                name: 'table9',
                ref: 'C' + row,
                headerRow: true,
                totalsRow: true,
                style: {
                    theme: 'TableStyleLight12',
                    showRowStripes: true,
                },
                columns: [
                    { name: 'Date' },
                    { name: 'Delivered', totalsRowFunction: 'sum' },
                    { name: 'Rejected', totalsRowFunction: 'sum' },
                    { name: 'Failed', totalsRowFunction: 'sum' },
                    { name: 'Total Messages', totalsRowFunction: 'sum' },
                    { name: 'Total Credit Used', totalsRowFunction: 'sum' },
                ],
                rows: [],
            });

            // Table Heading

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((el) => {
                msg91Worksheet.getCell(el).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'ADD8E6' },
                    bgColor: { argb: 'ADD8E6' },
                };
                msg91Worksheet.getCell(el).alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // Add Table

            let table9 = msg91Worksheet.getTable('table9');

            skillPassUpdateContactNoData.forEach((el: any) => {
                let date = Object.keys(el)[0];

                table9.addRow([
                    new Date(date),
                    el[date].Delivered,
                    el[date].Rejected,
                    el[date].Failed,
                    el[date].Total,
                    el[date].Credit,
                ]);
                ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                    msg91Worksheet.getCell(key).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' },
                    };
                    msg91Worksheet.getCell(key).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                    };
                });
                row++;
            });

            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            // Total Rows

            msg91Worksheet.getRow(row).font = { bold: true };
            ['C' + row, 'D' + row, 'E' + row, 'F' + row, 'G' + row, 'H' + row].map((key) => {
                msg91Worksheet.getCell(key).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
                msg91Worksheet.getCell(key).alignment = {
                    vertical: 'middle',
                    horizontal: 'center',
                };
            });
            row++;

            table9.commit();
            row++;
            row++;
            row++;

            await workbook.xlsx.writeFile('src/report/Output/output.xlsx');

            return createResponse(
                res,
                STATUS_CODE.OK,
                'Matrix Report Sheet is generated successfully.'
            );
        } catch (error: any) {
            createResponse(res, STATUS_CODE.INTERNAL_SERVER_ERROR, error);
        }
    }
}

export default new GenerateReport();
