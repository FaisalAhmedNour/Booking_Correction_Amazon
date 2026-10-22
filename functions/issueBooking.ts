import { AppCreator } from "../utils/App";
import { formatBoxType, generateInternalReferenceNo, isEmpty, sleep, twoDecimalNumber } from "../utils/common";
import { CenturyAPIProvider } from "../utils/century.api.provider";
import { IDataFormExcel } from "../utils/types";
import { Page } from "puppeteer-core";
import { CsvWriter } from "./po.upload.template.creator";
import { writeFileSync } from "fs";
import path from "path";

function isoToMMDDYYYY(isoStr: string): string {
    const date = new Date(isoStr);
    return date.toLocaleDateString("en-US", {
        month: "2-digit",
        day: "2-digit",
        year: "numeric",
    });
}
async function setInputFieldValue(
    key: string,
    val: string | undefined | null,
    selector: string,
    page: Page,
    errors: string[]
): Promise<boolean | void> {
    const val_ = String(val ?? '').trim();
    if (!val_) {
        return; // Empty optional input, skip
    }

    try {
        const res = await page.evaluate(
            (selector, value) => {
                const input = document.querySelector(selector);
                if (!input) return { success: false, message: 'eleNotFound' };
                // eslint-disable-next-line @typescript-eslint/ban-ts-comment
                // @ts-ignore
                input.focus();
                // eslint-disable-next-line @typescript-eslint/ban-ts-comment
                // @ts-ignore
                input.value = value;

                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('keydown', { bubbles: true }));
                input.dispatchEvent(new Event('keyup', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                input.dispatchEvent(new Event('blur', { bubbles: true }));

                return { success: true, message: 'ok' };
            },
            selector,
            val_
        );

        if (res.success === true) {
            return true;
        } else {
            if (res.message === 'eleNotFound') {
                errors.push(`Can't find input field "${key}" in the form.`);
            } else {
                errors.push(`Unknown issue #CFDED4 while setting value for "${key}".`);
            }
        }
    } catch (error) {
        const msg =
            error instanceof Error
                ? error.message
                : typeof error === 'string'
                    ? error
                    : 'Unknown error';

        errors.push(`Exception #E09D16s while setting input field "${key}": ${msg}`);
    }
}
async function setDateInputFieldValue(
    key: string,
    val: string | undefined | null,
    selector: string,
    page: Page,
    errors: string[]
): Promise<boolean | void> {
    const raw = String(val ?? '').trim();
    if (!raw) {
        errors.push(`Date field "${key}" is empty.`);
        return false;
    }

    // Parse YYYY-MM-DD or MM/DD/YYYY or Date string
    let year: number = 0, month: number = 0, day: number = 0;
    const isoMatch = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    const usMatch = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);

    if (isoMatch) {
        year = Number(isoMatch[1]);
        month = Number(isoMatch[2]);
        day = Number(isoMatch[3]);
    } else if (usMatch) {
        month = Number(usMatch[1]);
        day = Number(usMatch[2]);
        year = Number(usMatch[3]);
    } else {
        const parsed = new Date(raw);
        if (!isNaN(parsed.getTime())) {
            year = parsed.getFullYear();
            month = parsed.getMonth() + 1;
            day = parsed.getDate();
        } else {
            errors.push(`Invalid date format "${raw}" for "${key}". Please use MM/DD/YYYY or YYYY-MM-DD.`);
            return false;
        }
    }

    const formattedDate = `${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}/${year}`;

    try {
        const res = await page.evaluate(
            (selector, formattedDate, year, month, day) => {
                const input = document.querySelector(selector) as HTMLInputElement | null;

                if (!input) {
                    return { success: false, message: 'eleNotFound' };
                }

                input.focus();

                // Datepicker input
                if (input.classList.contains('hasDatepicker')) {
                    // @ts-ignore
                    if (typeof $ !== 'undefined' && $.fn?.datepicker) {
                        // @ts-ignore
                        $(input).datepicker(
                            'setDate',
                            new Date(year, month - 1, day)
                        );

                        input.dispatchEvent(
                            new Event('change', { bubbles: true })
                        );

                        input.dispatchEvent(
                            new Event('blur', { bubbles: true })
                        );

                        return { success: true, message: 'ok' };
                    }
                }

                // Normal input
                input.value = formattedDate;

                input.dispatchEvent(
                    new Event('input', { bubbles: true })
                );

                input.dispatchEvent(
                    new Event('keydown', { bubbles: true })
                );

                input.dispatchEvent(
                    new Event('keyup', { bubbles: true })
                );

                input.dispatchEvent(
                    new Event('change', { bubbles: true })
                );

                input.dispatchEvent(
                    new Event('blur', { bubbles: true })
                );

                return { success: true, message: 'ok' };
            },
            selector,
            formattedDate,
            year,
            month,
            day
        );

        if (res.success === true) {
            return true;
        }

        if (res.message === 'eleNotFound') {
            errors.push(`Can't find input field "${key}" in the form.`);
        } else if (res.message === 'invalidDate') {
            errors.push(`Invalid ISO date "${formattedDate}" for "${key}".`);
        } else {
            errors.push(`Could not set date "${formattedDate}" for "${key}".`);
        }
    } catch (error) {
        const msg =
            error instanceof Error
                ? error.message
                : typeof error === 'string'
                    ? error
                    : 'Unknown error';

        errors.push(
            `Exception #E09D16s while setting input field "${key}": ${msg}`
        );
    }
}
async function setSelectFieldValue(key: string, val: string | undefined | null, selector: string, page: Page, errors: string[]) {
    const rawVal = String(val ?? '').trim();
    if (!rawVal) {
        errors.push(`Select field "${key}" has an empty value.`);
        return false;
    }

    const normalize = (str: string) => str.replace(/\s+/g, ' ').trim().toLowerCase();
    const targetText = normalize(rawVal);
    let timeout = 30;
    while (timeout > 0) {

        try {
            const res = await page.evaluate((selector, targetText) => {
                const select = document.querySelector(selector) as HTMLSelectElement | null;
                if (!select) return { success: false, message: 'eleNotFound' };

                const normalizeInBrowser = (str: string) => (str || '').replace(/\s+/g, ' ').trim().toLowerCase();

                // Exact text match (case-insensitive, normalized whitespace)
                const option = Array.from(select.options).find(opt => normalizeInBrowser(opt.textContent || '') === targetText);

                if (option) {
                    select.value = option.value;
                    select.dispatchEvent(new Event('input', { bubbles: true }));
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    // @ts-ignore
                    if (typeof $ !== 'undefined') {
                        // @ts-ignore
                        $(select).trigger('change');
                    }
                    return { success: true, message: 'ok' }
                } else {
                    return { success: false, message: 'optNotFound' };
                }
            }, selector, targetText);

            if (res.success === true) {
                return true;
            }

            timeout--;
            if (timeout <= 0) {
                if (res.message === 'eleNotFound') errors.push(`Can't find ${key} select field in the form.`);
                else if (res.message === 'optNotFound') errors.push(`No option matching "${rawVal}" found in dropdown "${key}".`);
                else errors.push(`Unexpected error while selecting value for "${key}".`);
                return false;
            }
        } catch (error) {
            timeout--;
            console.info(key, 'retry', 30 - timeout, error);
            if (timeout <= 0) {
                errors.push(`Exception while selecting value for "${key}": ${error instanceof Error ? error.message : String(error)}`);
                return false;
            }
        }
        await sleep(1000);
    }
    errors.push(`Unexpected error #721B67 occurred while selecting value for "${key}".`);
    return false
}

async function getBookingNumber(app: AppCreator) {
    let bookingNumber = "";

    const startTime = Date.now();
    const timeout = 30_000;

    while (Date.now() - startTime < timeout) {
        const bookingText = await app.Page.$(
            '#page-content > quicknavigation-component > div > div > table > tbody > tr > td:nth-child(3) > span'
        );

        if (bookingText) {
            const innerText = await bookingText.evaluate(
                el => el.textContent?.trim() ?? ""
            );

            if (innerText.includes(":")) {
                const value = innerText.split(":")[1]?.trim();

                if (!isEmpty(value)) {
                    bookingNumber = value.trim().split(' ')[0];
                    return bookingNumber;
                }
            }
        }

        await sleep(2_000);
    }

    return null;
}

async function getBookingStatusFromPage(app: AppCreator): Promise<string | null> {
    try {
        const startTime = Date.now();
        const timeout = 10_000;

        while (Date.now() - startTime < timeout) {
            const status = await app.Page.evaluate(() => {
                // Check form status input/select or label
                const statusElem = document.querySelector('#VobHeaderForm_booking_status_desc') as HTMLInputElement | HTMLSelectElement | null;
                if (statusElem) {
                    if ('value' in statusElem && statusElem.value) {
                        return statusElem.value.trim();
                    }
                    if (statusElem.textContent) {
                        return statusElem.textContent.trim();
                    }
                }
                return null;
            });

            if (status && !isEmpty(status)) {
                return status;
            }

            await sleep(1_000);
        }
    } catch (e) {
        console.log("[getBookingStatusFromPage error]", e);
    }
}

async function getBookingCreatedDateFromPage(app: AppCreator): Promise<string | null> {
    try {
        const startTime = Date.now();
        const timeout = 10_000;

        while (Date.now() - startTime < timeout) {
            const createdDate = await app.Page.evaluate(() => {
                const elem = document.querySelector('#VobHeaderForm_create_date') as HTMLInputElement | null;
                if (elem) {
                    if ('value' in elem && elem.value && elem.value.trim()) {
                        return elem.value.trim();
                    }
                    if (elem.textContent && elem.textContent.trim()) {
                        return elem.textContent.trim();
                    }
                }
                const inputElem = document.querySelector('#VobHeaderForm_create_date input') as HTMLInputElement | null;
                if (inputElem) {
                    if (inputElem.value && inputElem.value.trim()) {
                        return inputElem.value.trim();
                    }
                    if (inputElem.textContent && inputElem.textContent.trim()) {
                        return inputElem.textContent.trim();
                    }
                }
                return null;
            });

            if (createdDate && !isEmpty(createdDate)) {
                return createdDate;
            }

            await sleep(1_000);
        }
    } catch (e) {
        console.log("[getBookingCreatedDateFromPage error]", e);
    }
    return null;
}

async function unitBookingHeader(app: AppCreator, data: IDataFormExcel[]) {
    try {
        const errors: string[] = []
        const commonData = app.initData.settings;
        const firstData = data[0]
        if (!commonData || !commonData.companyCode) {
            return {
                success: false,
                ref: "#E10001",
                message: "Company Code is missing in settings."
            };
        }

        const companyCode = commonData.companyCode.split('-')[0].trim()
        await app.Page.goto(`https://www.centuryvms.com/#/pagelayout?modulecd=Vob&configName=VobWorkspace&templateId=&supportingStatus=AddNewFromBlank&compCd=${companyCode}`, { timeout: 60_000 })
        // await app.Page.goto(`https://www.centuryvms.com/#/pagelayout?modulecd=Vob&configName=VobWorkspace&bookingKey=ACC589N01471BD0&compCd=${companyCode}`,{timeout:60_000})
        try {
            await app.Page.waitForSelector(`#page-content-button-list  button[title="Save Header Info"]`, {
                visible: true,
                timeout: 30_000
            });
        } catch (error) {
            await app.Page.reload()
            await app.Page.waitForSelector(`#page-content-button-list  button[title="Save Header Info"]`, {
                visible: true,
                timeout: 30_000
            });
        }

        await setSelectFieldValue("Ship/Vendor Code", commonData.shipVendorCode, "#VobHeaderForm_vendorsubvendorcode", app.Page, errors);

        await setInputFieldValue("Contact Phone (optional)", commonData.contactPhone as string, "#VobHeaderForm_phone", app.Page, errors);

        await setInputFieldValue("Contact Email", commonData.contactEmail, "#VobHeaderForm_email", app.Page, errors);


        await setSelectFieldValue("Country/Region of Export", "Bangladesh", "#VobHeaderForm_country_of_origin", app.Page, errors);

        await setSelectFieldValue("Add/Edit Manufacturer", commonData.addEditManufacturer, "#VobHeaderForm_manufacturer_cd", app.Page, errors);




        await setSelectFieldValue("Ship Mode", firstData.ship_mode, "#VobHeaderForm_ship_mode", app.Page, errors);
        await setSelectFieldValue("Freight Type", "CFS - to Century CFS facility", "#VobHeaderForm_freight_type", app.Page, errors);
        await setSelectFieldValue("Incoterms", firstData.incoterms, "#VobHeaderForm_incoterms", app.Page, errors);
        await setSelectFieldValue("FOB Port", firstData.fob_port, "#VobHeaderForm_fob_port", app.Page, errors);
        await setDateInputFieldValue("Estimated Cargo Delivery Date(MM/DD/YYYY)", firstData.ecdd, '#VobHeaderForm_est_rcvd_date input', app.Page, errors)

        console.error(errors)


        if (errors.length > 0) {
            return {
                success: false,
                ref: "#41194180",
                message: errors.join(', ')
            }
        }

        // Test Mode: Stop before submitting / saving
        if (commonData.reviewMode) {
            app.LocalCom.progress.setTitleAndMessage("Test Mode: Form Populated", "Header form filled. Stopped before submitting.");
            await app.LocalCom.sendData(undefined, [
                {
                    channel: app.LocalCom.progress.channel,
                    type: 'n',
                    title: `Test Mode: ${firstData.po || 'PO'}`,
                    message: `All form fields for PO "${firstData.po || ''}" (Model: "${firstData.model || ''}") are populated in Century VMS. Stopped before submitting/saving.`
                }
            ]);
            return {
                success: true,
                ref: "#TEST_MODE_PAUSED",
                message: "Test Mode: Form fields filled on screen (stopped before submitting).",
                isTestMode: true,
                data: {
                    companyCode,
                    bookingKey: "TEST_MODE_UNSAVED"
                }
            };
        }

        await app.Page.click(`#page-content-button-list  button[title="Save Header Info"]`);
        await app.Page.click(`#page-content-button-list  button[title="Save Header Info"]`);

        await sleep(1_000)
        const bookingNumber = await getBookingNumber(app);
        await sleep(1_000)

        await setInputFieldValue("Shipper Name", commonData.shipperName, "#VobHeaderForm_FCRShipperName", app.Page, errors);
        await setInputFieldValue("Email address to receive FCR", commonData.contactEmail, "#VobHeaderForm_FCREmailTo", app.Page, errors);

        await app.Page.click(`#page-content-button-list  button[title="Save Header Info"]`);
        await app.Page.click(`#page-content-button-list  button[title="Save Header Info"]`);

        const errorLabels = await app.Page.evaluate(() => {
            const xpath = '//label[following-sibling::validation-message-component/div]';

            const result = document.evaluate(
                xpath,
                document,
                null,
                XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,
                null
            );

            const labels = [];

            for (let i = 0; i < result.snapshotLength; i++) {
                const element = result.snapshotItem(i);
                labels.push(element?.textContent?.trim() ?? '');
            }

            return labels;
        });

        if (errorLabels.length > 0) {
            return {
                success: false,
                ref: "#98013267",
                message: errorLabels.join(', ') + " are invalid."
            }
        } else {
            if (bookingNumber) {
                return {
                    success: true,
                    ref: "#76866188",
                    message: "OK",
                    data: {
                        companyCode,
                        bookingKey: bookingNumber
                    }
                }
            } else {
                return {
                    success: false,
                    ref: "#91556597",
                    message: "Could not obtain booking key. #91556597"
                }
            }
        }


    } catch (error) {
        if (error && typeof error === 'object' && "known" in error && error.known) {
            return {
                success: false,
                ref: "ref" in error ? error.ref : "#71885168",
                message: "message" in error ? error.message : "#43446865"
            }
        } else {
            console.log(error)
            return {
                success: false,
                ref: '#0E0F00',
                message: "Unknown Error during Booking process."
            }
        }

    }
}
async function selectUnit(app: AppCreator, value: string) {
    await app.Page.waitForSelector('li[id="#VobItemDetail"]', { timeout: 20_000 })
    await app.Page.click('li[id="#VobItemDetail"]');
    // await app.Page.evaluate(() => {
    //     const tab = document.querySelector('li[id="#VobItemDetail"]') as HTMLElement | null;
    //     if (tab) tab.click();
    // });
    await app.Page.waitForSelector('div.ag-body-container div[role=gridcell][colid=po][tabindex="-1"]', { timeout: 20_000 })
    await app.Page.waitForSelector("div[colid=rowSelectCheckbox].ag-header-cell > span.ag-header-select-all", { timeout: 20_000 })
    await sleep(1_000);
    await app.Page.evaluate(() => {
        const selectAll = document.querySelector("div[colid=rowSelectCheckbox].ag-header-cell > span.ag-header-select-all, div[colid=rowSelectCheckbox].ag-header-cell input, div[colid=rowSelectCheckbox].ag-header-cell") as HTMLElement | null;
        if (selectAll) selectAll.click();
    });
    await app.Page.waitForSelector('div.ag-body-container div[role=gridcell][colid=carton_uom][tabindex="-1"]', { timeout: 20_000 })
    await app.Page.evaluate(() => {
        const cell = document.querySelector('div.ag-body-container div[role=gridcell][colid=carton_uom][tabindex="-1"]') as HTMLElement | null;
        if (cell) cell.click();
    });
    await sleep(1_000);
    const allOptions = await app.Page.$$("div.ui-menu-item-wrapper");

    const targetVal = value.trim().toLowerCase();
    for (const opt of allOptions) {
        const innerText = await opt.evaluate(
            el => (el.textContent || "").trim().toLowerCase()
        );

        if (innerText === targetVal || innerText.includes(targetVal)) {
            await opt.evaluate(el => (el as HTMLElement).click());
            break;
        }
    }
    await sleep(1_000);
    await app.Page.evaluate(() => {
        const updateBtn = document.querySelector("button[title=Update]") as HTMLElement | null;
        if (updateBtn) updateBtn.click();
    });
    try {
        await app.Page.waitForSelector('div[aria-label*="Successful"], div[aria-label*="Update"], .toast-success', { timeout: 30_000 });
    } catch {
        // Fallback in case toast dismissed quickly or selector variant is used
        await sleep(2_000);
    }
    return true;
}

async function validateFailedPoItems(apiProvider: CenturyAPIProvider, bookingKey: string, data: IDataFormExcel[]): Promise<string[]> {
    const invalidReasons: string[] = [];
    for (let i = 0; i < data.length; i++) {
        const item = data[i];
        const rowNum = i + 1;
        const po = (item.po || '').trim();
        const asin = (item.asin || '').trim();
        if (!po) {
            invalidReasons.push(`Row ${rowNum}: PO # is missing.`);
            continue;
        }
        if (!asin) {
            invalidReasons.push(`Row ${rowNum} (PO: ${po}): ASIN # is missing.`);
            continue;
        }
        if (!item.po_quantity || isNaN(Number(item.po_quantity)) || Number(item.po_quantity) <= 0) {
            invalidReasons.push(`Row ${rowNum} (PO: ${po}, ASIN: ${asin}): Invalid Booked Units / PO Quantity "${item.po_quantity}".`);
        }
        if (!item.carton_qty || isNaN(Number(item.carton_qty)) || Number(item.carton_qty) <= 0) {
            invalidReasons.push(`Row ${rowNum} (PO: ${po}, ASIN: ${asin}): Invalid Carton Quantity "${item.carton_qty}".`);
        }
        if (item.net_weight && isNaN(Number(item.net_weight))) {
            invalidReasons.push(`Row ${rowNum} (PO: ${po}, ASIN: ${asin}): Net Weight "${item.net_weight}" is not a valid number.`);
        }
        if (item.gross_weight && isNaN(Number(item.gross_weight))) {
            invalidReasons.push(`Row ${rowNum} (PO: ${po}, ASIN: ${asin}): Gross Weight "${item.gross_weight}" is not a valid number.`);
        }
        try {
            const searchRes = await apiProvider.searchBookingItem(bookingKey, po, asin);
            if (!searchRes.found) {
                invalidReasons.push(`Row ${rowNum} (PO: ${po}, ASIN: ${asin}): Not found in Century VMS for booking "${bookingKey}". ${searchRes.message || ''}`.trim());
            }
        } catch (apiErr: any) {
            invalidReasons.push(`Row ${rowNum} (PO: ${po}, ASIN: ${asin}): Century VMS API search error - ${apiErr?.message || String(apiErr)}`);
        }
    }
    return invalidReasons;
}

async function unitBookingUploadPO(app: AppCreator, apiProvider: CenturyAPIProvider, data: IDataFormExcel[], payload: {
    companyId: string,
    bookingKey: string,
}) {
    console.log('79922373')
    try {
        await app.Page.goto(`https://www.centuryvms.com/#/pagelayout?modulecd=Vob&configName=VobWorkspace&bookingKey=${payload.bookingKey}&compCd=${payload.companyId}`, { waitUntil: 'load' })
        // await app.Page.goto(`https://www.centuryvms.com/#/pagelayout?modulecd=Vob&configName=VobWorkspace&bookingKey=ACC589N01471BD0&compCd=${payload.companyId}`,{waitUntil:'load'})
        await app.Page.reload({ waitUntil: "load" })
        // const bookingNumber = await getBookingNumber(app)
        // console.log("47801003", bookingNumber)
        const errors: string[] = []
        try {
            console.log('15644736')

            await app.Page.waitForSelector('li[id="#VobItemDetail"]', { timeout: 20_000 })
            await app.Page.evaluate(() => {
                const tab = document.querySelector('li[id="#VobItemDetail"]') as HTMLElement | null;
                if (tab) tab.click();
            });
        } catch (error) {
            console.log("45317763", error)
            return {
                success: false,
                ref: "#89547106",
                message: "Error in saving headers."
            }
        }

        try {
            await app.Page.waitForSelector('grid-component .spinner-container, span.ag-overlay-loading-center', {
                hidden: true,
                timeout: 15_000
            });
        } catch {
            // Spinner already gone or didn't appear
        }

        const hasExistingRows = await app.Page.evaluate(() => {
            const rows = document.querySelectorAll(
                'div.ag-body-container div[role="row"], div.ag-center-cols-container div.ag-row, div.ag-body-container div.ag-row'
            );
            return rows.length > 0;
        });

        if (hasExistingRows) {
            console.log('[unitBookingUploadPO] Existing items found. Deleting all...');
            await app.Page.evaluate(() => {
                const selectAll = document.querySelector('div[colid=rowSelectCheckbox].ag-header-cell > span.ag-header-select-all, div[colid=rowSelectCheckbox].ag-header-cell input, div[colid=rowSelectCheckbox].ag-header-cell') as HTMLElement | null;
                if (selectAll) selectAll.click();
            });
            await sleep(1_000);

            await app.Page.evaluate(() => {
                const deleteBtn = document.querySelector('button[title=Delete], button[title="Delete"]') as HTMLElement | null;
                if (deleteBtn && !deleteBtn.hasAttribute('disabled')) deleteBtn.click();
            });

            try {
                const confirmBtn = await app.Page.waitForSelector(
                    '.modal-dialog button[title="Continue"], .modal-dialog button[title="Yes"], .modal-dialog button[title="OK"], .modal-footer button.btn-primary',
                    { visible: true, timeout: 2_000 }
                );
                if (confirmBtn) {
                    await app.Page.evaluate(() => {
                        const btn = document.querySelector('.modal-dialog button[title="Continue"], .modal-dialog button[title="Yes"], .modal-dialog button[title="OK"], .modal-footer button.btn-primary') as HTMLElement | null;
                        if (btn) btn.click();
                    });
                }
            } catch {
                // No confirmation popup, continue
            }

            try {
                await app.Page.waitForSelector('grid-component .spinner-container, span.ag-overlay-loading-center', {
                    hidden: true,
                    timeout: 30_000
                });
            } catch {
                // Deletion was instant
            }
            await sleep(500);
        } else {
            console.log('[unitBookingUploadPO] Table is empty, skipping delete.');
        }


        const excelTemplateMaker = await CsvWriter.create();
        for (const current_data of data) {

            await excelTemplateMaker.appendRow([
                ,                   // "Serial Number",
                current_data.po,             // "PO #",
                current_data.asin,           // "ASIN #",
                current_data.model,          // "Model #",
                ,                   // "ASIN FC",
                ,                   // "Add/Edit Country/Region of Origin (COO)",
                ,                   // "Ordered Units",
                current_data.po_quantity,     // "Booked Units",
                ,                   // "Unit of Measure",
                ,                   // "Booked Units in other booking",
                current_data.carton_qty,   // "Booked Cartons",
                current_data.net_weight,  // "Net Weight (KG)",
                current_data.cbm,    // "Booked Volume",
                ,                   // "Volume UOM",
                current_data.gross_weight,    // "Booked Weight",
                ,                   // "Weight UOM",
                ,                   // "Last Ship Date",
                ,                   // "Estimated Delivery Date",
                ,                   // "PO Item Description",
                ,                   // "Carrier Item Description",
                ,                   // "Stuffing Location Name",
                ,                   // "Manufacturer Name",
                ,                   // "Shipping Marks",
                ,                   // "Harmonized Tariff Code",
                ,                   // "Inspection PO Approve",
                ,                   // "Inspection Approve",
                ,                   // "Cargo Type ID",
                ,                   // "Cargo Type",
                ,                   // "Item Certificate Approved",
                ,                   // "Remarks in FCR",
                ,                   //"Sort"
            ])
        }

        const csvBuffer = await excelTemplateMaker.getBuffer();
        const fileName = payload.bookingKey;
        writeFileSync(path.join(path.dirname(app.initData.excelFile), fileName + "_PO_items.csv"), new Uint8Array(csvBuffer));

        const uploadRes = await apiProvider.uploadCSV(csvBuffer, payload.bookingKey, payload.companyId);
        await excelTemplateMaker.cleanup()

        if (!uploadRes.success) {
            console.log("[unitBookingUploadPO] CSV upload failed, running item validation against Century VMS API...");
            const itemErrors = await validateFailedPoItems(apiProvider, payload.bookingKey, data);
            let combinedMsg = uploadRes.message || "Failed to upload PO items CSV.";
            if (itemErrors.length > 0) {
                combinedMsg = `PO Upload failed. Invalid item(s) found:\n${itemErrors.join('\n')}`;
            }
            return {
                success: false,
                ref: "#PO_UPLOAD_INVALID_DATA",
                message: combinedMsg
            };
        }

        await app.Page.reload()

        await selectUnit(app, "PIECES")
        try {
            const bs64 = await apiProvider.downloadPDF(app.Page, payload.bookingKey);
            if (bs64) {
                const pdfBuffer = Buffer.from(bs64, "base64");
                writeFileSync(path.join(path.dirname(app.initData.excelFile), fileName + ".pdf"), new Uint8Array(pdfBuffer))
            }
        } catch (pdfError) {
            console.log("[downloadPDF error]", pdfError);
        }
        return true
    } catch (error) {
        if (error && typeof error === 'object' && "known" in error && error.known) {
            return {
                success: false,
                ref: "ref" in error ? error.ref : "#71885168",
                message: "message" in error ? error.message : "#43446865"
            }
        } else {
            console.log(error)
            return {
                success: false,
                ref: '#0E0F00',
                message: error instanceof Error ? error.message : "Unknown Error during Booking process."
            }
        }

    }
}

async function submitBooking(app: AppCreator, payload: {
    companyId: string
    bookingKey: string
}) {
    try {
        const targetUrl = `https://www.centuryvms.com/#/pagelayout?modulecd=Vob&configName=VobWorkspace&bookingKey=${payload.bookingKey}&compCd=${payload.companyId}`;
        const currentUrl = app.Page.url();
        if (!currentUrl.includes(`bookingKey=${payload.bookingKey}`)) {
            await app.Page.goto(targetUrl, { waitUntil: 'load' });
        }

        // Wait for the Send button in the action bar to become visible
        try {
            await app.Page.waitForSelector('#page-content-button-list button[title="Send"]', {
                visible: true,
                timeout: 30_000
            });
        } catch (waitErr) {
            console.log("[submitBooking] Send button not visible initially, reloading page...", waitErr);
            await app.Page.reload({ waitUntil: "load" });
            await app.Page.waitForSelector('#page-content-button-list button[title="Send"]', {
                visible: true,
                timeout: 30_000
            });
        }

        await sleep(1_000);

        // Click Send button
        await app.Page.evaluate(() => {
            const sendBtn = document.querySelector('#page-content-button-list button[title="Send"]') as HTMLButtonElement | null;
            if (sendBtn) {
                sendBtn.click();
            }
        });

        // Handle confirmation modal if it appears (e.g. "Continue")
        try {
            await app.Page.waitForSelector('.modal-dialog button[title="Continue"], .modal-body button[title="Continue"], button[title="Continue"]', {
                visible: true,
                timeout: 15_000
            });
            await app.Page.evaluate(() => {
                const continueBtn = (
                    document.querySelector('.modal-dialog button[title="Continue"]') ||
                    document.querySelector('.modal-body button[title="Continue"]') ||
                    document.querySelector('button[title="Continue"]')
                ) as HTMLButtonElement | null;
                if (continueBtn) {
                    continueBtn.click();
                }
            });
        } catch (modalErr) {
            console.log("[submitBooking] Continue modal not shown or already dismissed:", modalErr);
        }

        await sleep(3_000);
        const liveStatus = await getBookingStatusFromPage(app);
        return {
            success: true,
            status: liveStatus || "Sent",
            ref: "#SUCCESS_SUBMIT",
            message: "Booking submitted successfully."
        };
    } catch (error) {
        if (error && typeof error === 'object' && "known" in error && error.known) {
            return {
                success: false,
                ref: "ref" in error ? error.ref : "#71885168",
                message: "message" in error ? error.message : "#43446865"
            }
        } else {
            console.log(error)
            return {
                success: false,
                ref: '#0E0F00',
                message: error instanceof Error ? error.message : "Unknown Error during Booking process."
            }
        }
    }
}

async function openAndVerifyExistingBooking(app: AppCreator, bookingKey: string, data: IDataFormExcel[]) {
    try {
        const errors: string[] = [];
        const commonData = app.initData.settings;
        const firstData = data[0];
        if (!commonData || !commonData.companyCode) {
            return {
                success: false,
                ref: "#E10001",
                message: "Company Code is missing in settings."
            };
        }

        const companyCode = commonData.companyCode.split('-')[0].trim();
        // await app.Page.goto(`https://www.centuryvms.com/#/pagelayout?modulecd=Vob&configName=VobWorkspace&bookingKey=ACC589N01471BD0&compCd=${companyCode}`,{timeout:60_000})
        const targetUrl = `https://www.centuryvms.com/#/pagelayout?modulecd=Vob&configName=VobWorkspace&bookingKey=${encodeURIComponent(bookingKey)}&compCd=${companyCode}`;
        console.log(`[openAndVerifyExistingBooking] Navigating to: ${targetUrl}`);

        await app.Page.goto(targetUrl, { waitUntil: 'load', timeout: 60_000 });

        // Wait for page action buttons or quick navigation to be visible
        try {
            await app.Page.waitForSelector('#page-content, #page-content-button-list, quicknavigation-component', {
                visible: true,
                timeout: 30_000
            });
        } catch (error) {
            console.log("[openAndVerifyExistingBooking] Initial wait timed out, reloading page...");
            await app.Page.reload({ waitUntil: 'load' });
            await app.Page.waitForSelector('#page-content, #page-content-button-list, quicknavigation-component', {
                visible: true,
                timeout: 30_000
            });
        }

        await sleep(2_000);

        // Verify booking loaded
        const liveBookingNumber = await getBookingNumber(app);
        console.log(`[openAndVerifyExistingBooking] Live booking detected: "${liveBookingNumber}", expected: "${bookingKey}"`);

        // Check if page displays an error message or alert
        const pageError = await app.Page.evaluate(() => {
            const errBanner = document.querySelector('.alert-danger, .toast-error, .validation-summary-errors');
            return errBanner ? errBanner.textContent?.trim() : null;
        }).catch(() => null);

        if (pageError && !liveBookingNumber) {
            return {
                success: false,
                ref: "#BOOKING_NOT_FOUND",
                message: `Century VMS Error for Booking "${bookingKey}": ${pageError}`
            };
        }

        // Test Mode: Stop before performing edits/upload
        if (commonData.reviewMode) {
            app.LocalCom.progress.setTitleAndMessage("Test Mode: Booking Loaded", `Booking ${bookingKey} opened. Stopped before applying corrections.`);
            await app.LocalCom.sendData(undefined, [
                {
                    channel: app.LocalCom.progress.channel,
                    type: 'n',
                    title: `Test Mode: Booking ${bookingKey}`,
                    message: `Booking "${bookingKey}" (PO: "${firstData.po || ''}") is open and verified in Century VMS. Stopped before applying corrections/deleting items.`
                }
            ]);
            return {
                success: true,
                ref: "#TEST_MODE_PAUSED",
                message: "Test Mode: Existing booking verified on screen (stopped before applying corrections).",
                isTestMode: true,
                data: {
                    companyCode,
                    bookingKey
                }
            };
        }

        return {
            success: true,
            ref: "#BOOKING_LOADED_OK",
            message: "OK",
            data: {
                companyCode,
                bookingKey: liveBookingNumber || bookingKey
            }
        };

    } catch (error) {
        if (error && typeof error === 'object' && "known" in error && (error as any).known) {
            return {
                success: false,
                ref: "ref" in error ? (error as any).ref : "#71885168",
                message: "message" in error ? (error as any).message : "#43446865"
            };
        } else {
            console.log(error);
            return {
                success: false,
                ref: '#0E0F00',
                message: error instanceof Error ? error.message : "Unknown error while loading existing booking."
            };
        }
    }
}


export default async function issueBooking(app: AppCreator) {
    await app.Page.goto("https://www.centuryvms.com/#/pagelayout?moduleCd=Vob&configName=VobMain", { timeout: 60_000 })
    await sleep(30_000)
    await app.Page.reload()
    const vmaApiProvider = new CenturyAPIProvider(await app.createSession())
    app?.LocalCom?.progress?.setTotal(app.dateManager.length)
    const sesResp = await app.LocalCom.progress.getNewSession(5, !app.LocalCom.CLIargs.devFlag)
    if (!sesResp.doNext) {
        app.LocalCom.progress.setTitleAndMessage("Stopped.", sesResp.message)
        await app.LocalCom.sendData(undefined, [
            {
                channel: app.LocalCom.progress.channel,
                type: 'e',
                title: 'Stopped.',
                message: sesResp.message || "#88388484"
            }
        ])
        return
    }

    app.LocalCom.progress.setTitleAndMessage('Booking', 'Booking is running.')
    app.LocalCom.sendData();
    while (app.dateManager.haveNext()) {
        let data: {
            Success: boolean,
            Message: string,
            data: IDataFormExcel[],
            booking_number?: string,
            booking_created?: string,
            po_status?: string,
        } = {
            Success: false,
            Message: "",
            data: [],
            booking_number: undefined,
            booking_created: undefined,
            po_status: undefined,
        }
        try {
            const _data = app.dateManager.getData()
            if (_data && _data.length > 0) data = {
                Success: false,
                Message: "",
                data: _data,
                booking_number: undefined,
                booking_created: undefined,
                po_status: undefined,
            };

            if (isEmpty(data.data) || data.data.length === 0) {
                data.Message = "No PO row data found for this group.";
            } else {
                const firstRow = data.data[0];
                const poNumber = firstRow.po || "Unknown PO";

                const bookingNumber = (firstRow.booking_number || (firstRow as any)['Booking Number'] || (firstRow as any)['Draft Booking No'] || '').trim();

                if (isEmpty(bookingNumber)) {
                    throw {
                        known: true,
                        ref: "#MISSING_BOOKING_NUMBER",
                        message: `Row group for PO "${poNumber}" is missing the required "Booking Number" in Excel.`
                    };
                }

                data.booking_number = bookingNumber;

                // Pre-flight validation of mandatory Excel fields
                const missingFields: string[] = [];
                if (isEmpty(firstRow.po)) missingFields.push("PO");
                if (isEmpty(firstRow.model)) missingFields.push("Model");
                if (isEmpty(firstRow.ship_mode)) missingFields.push("Ship Mode");
                if (isEmpty(firstRow.incoterms)) missingFields.push("Incoterms");
                if (isEmpty(firstRow.fob_port)) missingFields.push("FOB Port");
                if (isEmpty(firstRow.ecdd)) missingFields.push("ECDD");

                for (let idx = 0; idx < data.data.length; idx++) {
                    const row = data.data[idx];
                    if (isEmpty(row.asin)) missingFields.push(`ASIN (Row ${idx + 1})`);
                    if (isEmpty(row.po_quantity)) missingFields.push(`PO Qty (Row ${idx + 1})`);
                    if (isEmpty(row.carton_qty)) missingFields.push(`Carton Qty (Row ${idx + 1})`);
                }

                if (missingFields.length > 0) {
                    throw {
                        known: true,
                        ref: "#VALIDATION_ERROR",
                        message: `Missing required field(s) in Excel: ${missingFields.join(', ')}`
                    };
                }

                const verifyResp = await openAndVerifyExistingBooking(app, bookingNumber, data.data);
                console.log("verifyResp", verifyResp);

                if (verifyResp.success) {
                    data.booking_number = verifyResp.data?.bookingKey
                    const createdDate = await getBookingCreatedDateFromPage(app);
                    if (createdDate) {
                        data.booking_created = createdDate;
                    }
                    const fileUpload = await unitBookingUploadPO(app, vmaApiProvider, data.data, {
                        companyId: verifyResp.data?.companyCode as string,
                        bookingKey: verifyResp.data?.bookingKey as string
                    })
                    if (fileUpload === true) {
                        data.Success = true
                        data.Message = 'OK'
                    } else if (typeof fileUpload === 'object' && 'message' in fileUpload) {
                        data.Success = false;
                        data.Message = (fileUpload as any).message;
                    } else {
                        data.Success = false
                        data.Message = "Booking header was saved but failed to upload PO."
                    }
                } else {
                    throw {
                        known: true,
                        ref: verifyResp.ref,
                        message: verifyResp.message,
                    }
                }

                // Determine booking status and handle submission if requested
                let bookingStatus = "Draft";
                if ((verifyResp as any).isTestMode) {
                    bookingStatus = "Test Mode (Unsaved)";
                } else if (data.Success) {
                    // Check live status from Century VMS page
                    const liveStatus = await getBookingStatusFromPage(app);
                    if (liveStatus) {
                        bookingStatus = liveStatus;
                    }

                    // If user selected "Sent" status in settings and not in reviewMode, submit the booking
                    if (app.initData.settings.status === 'Sent' && !app.initData.settings.reviewMode) {
                        const submitRes = await submitBooking(app, {
                            companyId: verifyResp.data?.companyCode as string,
                            bookingKey: verifyResp.data?.bookingKey as string
                        });

                        if (submitRes && submitRes.success) {
                            bookingStatus = submitRes.status || "Sent";
                        } else if (submitRes && !submitRes.success) {
                            bookingStatus = "Submit Failed";
                            data.Success = false;
                            data.Message = submitRes.message || "Failed to submit booking.";
                        }
                    }
                } else {
                    bookingStatus = "Failed";
                }

                if (!data.booking_created) {
                    const liveCreatedDate = await getBookingCreatedDateFromPage(app);
                    if (liveCreatedDate) {
                        data.booking_created = liveCreatedDate;
                    }
                }

                data.po_status = bookingStatus;

                // Propagate booking status, booking number, and booking created date to all row records for UI table display
                if (Array.isArray(data.data)) {
                    for (const row of data.data) {
                        row.po_status = bookingStatus;
                        if (data.booking_number) {
                            row.booking_number = data.booking_number;
                        }
                        if (data.booking_created) {
                            row.booking_created = data.booking_created;
                        }
                    }
                }
            }
        } catch (e) {
            data.Success = false
            if (e && typeof e === 'object' && 'known' in e && (e as any).known === true) {
                data.Message = (e as any).message || '';
            } else if (e instanceof Error) {
                data.Message = e.message;
            } else {
                data.Message = String(e || 'Unknown error occurred.');
            }
            data.po_status = "Failed";
            if (Array.isArray(data.data)) {
                for (const row of data.data) {
                    row.po_status = "Failed";
                    if (data.booking_number) row.booking_number = data.booking_number;
                    if (data.booking_created) row.booking_created = data.booking_created;
                }
            }
        }
        finally {
            let resp: { continue: boolean; message: string | undefined; }
            const poTitle = (data.data && data.data[0] && data.data[0].po) ? data.data[0].po : "Unknown";
            if (data.Success) {

                resp = await app.LocalCom.progress.registerSuccess(1)
            } else {
                app.LocalCom.sendData(undefined, [{
                    channel: app.LocalCom.progress.channel,
                    type: 'e',
                    title: `PO: ${poTitle}`,
                    message: `PO: ${poTitle} failed. Because, \n ${data.Message}`
                }])
                resp = await app.LocalCom.progress.registerFail(1)
            }
            await app.LocalCom.sendData([data])
            if (!resp.continue) {
                app.LocalCom.progress.setTitleAndMessage("Stopped.", resp.message)
                app.LocalCom.sendData(undefined, [
                    {
                        channel: app.LocalCom.progress.channel,
                        type: 'e',
                        title: 'Stopped.',
                        message: resp.message || "#33017223"
                    }
                ])
                break
            }
            app.dateManager.next()
            // await sleep(3000000)
        }

    }

}