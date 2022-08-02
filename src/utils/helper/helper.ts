import { Response } from 'express';

/**
 * @description Create Response
 * @param {Object} res
 * @param {Number} status
 * @param {String} message
 * @param {Object} payload
 * @param {Object} pager
 */

export const createResponse = (
    res: Response,
    status: number,
    message: string,
    payload: object | null = {},
    pager: object | null = {}
) => {
    return res.status(status).json({
        status,
        message,
        payload,
        pager: typeof pager !== 'undefined' ? pager : {},
    });
};
