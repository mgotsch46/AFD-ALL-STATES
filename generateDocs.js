const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, UnderlineType, VerticalAlign
} = require('docx');
const archiver = require('archiver');

const ARIAL = 'Arial';
const LETTER = { width: 12240, height: 15840 };
const LEGAL  = { width: 12240, height: 20160 };
// Tight margins for dense docs: 0.6" top/bottom, 0.75" sides
const MARGIN      = { top: 1080, right: 1080, bottom: 1080, left: 1080 };
const MARGIN_TIGHT = { top: 864,  right: 1008, bottom: 864,  left: 1008 };
const FULL      = 10080; // content width with 0.75" margins on letter
const FULL_TIGHT = 10224; // content width with 0.7" side margins

const CHECK = '\u2714'; // ✔

function r(text, opts = {}) {
  return new TextRun({
    text: String(text || ''), font: ARIAL,
    size: opts.size || 20,
    bold: opts.bold || false,
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
    italics: opts.italics || false,
  });
}
function p(children, opts = {}) {
  if (typeof children === 'string') children = [r(children, { size: opts.size })];
  return new Paragraph({
    children,
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after !== undefined ? opts.after : 100, line: opts.line || 276 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    pageBreakBefore: opts.pageBreak || false,
    border: opts.border || undefined,
  });
}
function blank(sz) { return p('', { after: sz !== undefined ? sz : 60 }); }
function hRule() {
  return p('', { border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000', space: 1 } }, before: 40, after: 40 });
}
function heading(text, size) {
  return p([r(text, { bold: true, size: size || 22 })], { align: AlignmentType.CENTER, before: 80, after: 60 });
}
function subheading(text) {
  return p([r(text, { bold: true, size: 20 })], { align: AlignmentType.CENTER, before: 40, after: 60 });
}
function sectionHead(text, size) {
  return p([r(text, { bold: true, size: size || 20, underline: true })], { before: 80, after: 40 });
}
function nb() {
  const n = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top: n, bottom: n, left: n, right: n };
}
function bdr() {
  const b = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  return { top: b, bottom: b, left: b, right: b };
}
function cell(text, w, opts = {}) {
  return new TableCell({
    borders: opts.noborder ? nb() : bdr(),
    width: { size: w, type: WidthType.DXA },
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [p([r(text, { bold: opts.bold || false, size: opts.size || 19 })], { align: opts.align || AlignmentType.LEFT, after: 0 })]
  });
}

// ─── Agreement for Deed ───────────────────────────────────────────────────────
function buildAgreementForDeed(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const buyerStr = buyers.join(' and ');
  const children = [
    heading('AGREEMENT FOR DEED'),
    blank(40),
    p([r('THIS AGREEMENT FOR DEED is made and entered into this '), r(d.doc_date, {bold:true}), r(', by and between, '), r(d.seller_name, {bold:true}), r(' (hereinafter referred to as "First Party"), and '), r(buyerStr, {bold:true}), r(' (hereinafter referred to as "Second Party").')]),
    blank(40),
    p([r('WITNESSETH, that if the Second Party shall first make all of the payments and fully perform the covenants hereinafter mentioned on his part to be made and performed, the First Party hereby covenants and agrees to convey to the Second Party, his heirs, executors, administrators, personal representatives, or assigns, in fee simple absolute, clear of all encumbrances, except as set out herein, by a good and sufficient warranty deed, that parcel of land situated in the City of '), r(d.city, {bold:true}), r(', State of '), r(d.state, {bold:true}), r(', known as:')]),
    blank(20),
    p([r(d.property_address, {bold:true})], {align: AlignmentType.CENTER}),
    p([r('Legal Description: '), r(d.legal_description || '________________________________________', {bold:true})]),
    p([r('P.I.N. #: '), r(d.pid || '________________________________________', {bold:true})]),
    p([r('(the "Property")')], {align: AlignmentType.CENTER}),
    blank(40),
    p([r('1.  ', {bold:true}), r('The Second Party herein covenants and agrees to pay to the First Party the sum of '), r(d.purchase_price_words, {bold:true}), r(' ('), r(d.purchase_price_numbers, {bold:true}), r('), in the following manner:')]),
    blank(20),
    p([r('a.  '), r(d.down_payment_words, {bold:true}), r(' ('), r(d.down_payment_numbers, {bold:true}), r(') paid to the First Party. Down payment is to be made out to '), r(d.seller_name, {bold:true}), r('.')], {indent:360}),
    blank(20),
    p([r('b.  The principal sum of '), r(d.loan_amount_words, {bold:true}), r(' ('), r(d.loan_amount_numbers, {bold:true}), r('), with interest thereon at the rate of '), r(d.interest_rate_words, {bold:true}), r(' percent ('), r(d.interest_rate_numbers, {bold:true}), r('%) per annum, in monthly installments of '), r(d.pi_words, {bold:true}), r(' ('), r(d.pi_numbers, {bold:true}), r('), beginning on the '), r(d.first_payment_date, {bold:true}), r(', and continuing on the 1st day of each and every month thereafter until paid in full. Payments are considered late after the 5th of the month and will incur a 10% late fee. ALL Monthly Payments are to be made out to '), r(d.seller_name, {bold:true}), r(' and payable by: '), r(d.payment_method || 'Buildium', {bold:true}), r('. This Agreement is amortized from '), r(d.first_payment_date, {bold:true}), r(' to '), r(d.maturity_date, {bold:true}), r(', with final payment due on '), r(d.maturity_date, {bold:true}), r(', unless prepaid beforehand. Property Taxes for '), r(new Date().getFullYear().toString(), {bold:true}), r(' are '), r(d.monthly_taxes_numbers, {bold:true}), r(' per month. Property Taxes shall be reviewed yearly and amended accordingly. Loan servicing fee of '), r(d.servicing_fee || '$120.00', {bold:true}), r(' charged each month.')], {indent:360}),
    blank(20),
    p([r('2.  Conveyance shall be by special warranty deed upon completion of all payments. The second party will not receive any tax documents from the first party and shall rely solely on the amortization schedule for tax purposes. The First Party grants the right of possession and occupancy upon acceptance of this agreement.')]),
    blank(20),
    p([r('3.  Upon completion of all payments, title shall be conveyed free and clear of all encumbrances except easements, restrictions, limitations, reservations, covenants and conditions of record, applicable zoning ordinances, and real estate taxes for the year in which the deed is delivered.')]),
    blank(20),
    p([r('4.  The Second Party agrees to pay Full Replacement Insurance and is required to contact an agent within 72 hours to obtain a homeowners policy. The First Party shall be listed as an additional insured party. A default of insurance shall be considered a default of this Agreement.')]),
    blank(20),
    p([r('5.  This is a business transaction and the Second Party is not purchasing the property but paying the First Party for possession and use of the property, as well as the potential future right to become the owner of the Property.')]),
    blank(20),
    p([r('6.  In the event of eminent domain, the First Party will be paid the remaining principal balance and 100% of remaining proceeds will be paid to the Second Party.')]),
    blank(20),
    p([r('7.  The Second Party may prepay the FULL principal balance at any time without penalty. Upon full prepayment, the First Party shall have twenty (20) days to deliver a warranty deed. NO PARTIAL PRE-PAYMENTS UNLESS MUTUALLY AGREED UPON.')]),
    blank(20),
    p([r('8.  In case of failure by the Second Party to make payments or perform any covenants, this Agreement shall, at the option of the First Party, be forfeited and terminated, and all payments made shall be retained by the First Party as liquidated damages. The First Party shall have the right to demand immediate possession and commence eviction proceedings. All costs including reasonable attorneys\u2019 fees shall be paid by the Second Party.')]),
    blank(20),
    p([r('9.  The time of each payment is an essential part of this contract. Any payment not received within ten (10) days of its due date, or three payments not received within five (5) days of their due dates, shall constitute a default under paragraph 8.')]),
    blank(20),
    p([r('10. Notices may be sent by mail to the last known address of the party to be notified.')]),
    blank(20),
    p([r('11. All covenants shall extend to and be obligatory upon successors, heirs, executors, administrators, and assigns of the respective parties.')]),
    blank(20),
    p([r('12. The Second Party will not permit or commit waste to the Property and will maintain it in good repair at all times. Failure to maintain the Property shall be an event of default and the First Party may exercise all rights set forth in Paragraph 8.')]),
    blank(20),
    p([r('13. The words "First Party" and "Second Party" shall be construed to include the plural as well as the singular, and the masculine shall include the feminine and neuter where context requires.')]),
    blank(20),
    p([r('14. This agreement shall be governed by the laws of the State of '), r(d.state, {bold:true}), r('. The parties waive trial by jury and submit to jurisdiction in '), r(d.county, {bold:true}), r(' County, '), r(d.state, {bold:true}), r('. The prevailing party in any litigation shall be entitled to recover reasonable attorney\u2019s fees and costs.')]),
    blank(20),
    p([r('15. The First Party may sign and record a Notice of Termination if the Second Party defaults, effective if no legal proceeding or lis pendens is filed within three months of recordation.')]),
    blank(20),
    p([r('16. Service of process may be obtained through certified mail, return receipt requested; the parties hereto waiving any right to object to the method by which service was perfected.')]),
    blank(20),
    p([r('17. The First Party may inspect the property with 72 hours verbal notice every 6 months. If inspection shows damage or degradation, a 30-day notice to repair shall be given. Failure to repair shall constitute a default under Paragraph 8.')]),
    blank(20),
    p([r('18. The Second Party agrees to notify the First Party immediately upon discovering serious building problems such as foundation cracks, roof leaks, moisture, or termite activity.')]),
    blank(20),
    p([r('19. The Second Party shall not engage in or allow illegal activities on the premises.')]),
    blank(20),
    p([r('20. The Second Party shall keep the property insured for full replacement cost with First Party listed as additional insured until the property has been paid in full.')]),
    blank(20),
    p([r('21. The Second Party agrees to assume the Alarm System and maintain monitoring services for the duration of this agreement.')]),
    blank(20),
    p([r('22. The First Party shall not be responsible for any repairs. All repairs and maintenance are the Second Party\u2019s sole responsibility and cost.')]),
    blank(20),
    p([r('23. The Second Party agrees to transfer all utilities to their names within 48 hours after move-in.')]),
    blank(20),
    p([r('24. The Second Party acknowledges they have read this Agreement, understand it, agree to it, and have been given a copy. They have been advised to seek legal, tax, and technical counsel prior to signing.')]),
    blank(40),
    hRule(),
    p([r('DISCLAIMER: ', {bold:true}), r('By signing below, you acknowledge that this transaction is not a traditional sale and purchase of real property. You (the "Second Party") are not purchasing the property but paying the "First Party" for possession and use of the property, as well as the potential future right to become the owner. YOU WILL NOT OWN THIS PROPERTY UPON SIGNING THIS AGREEMENT OR TAKING POSSESSION. You will not take title until you fulfill all terms including all timely payments. If you breach any obligation, '), r(d.seller_name, {bold:true}), r(' will have the right to terminate this Agreement and treat you as a tenant, subject to potential eviction. You are strongly encouraged to seek independent legal advice.')]),
    blank(20),
    p([r('Total monthly payment is '), r(d.total_monthly_payment || '', {bold:true}), r(', this includes property taxes.')]),
    blank(20),
    p([r('The Parties have hereunto set their hands and seals the day and year first above written.')]),
    blank(40),
    p([r('"FIRST PARTY"', {bold:true})]),
    blank(20),
    p([r('________________________________________')]),
    p([r(d.seller_name)]),
    p([r('Signature')]),
    blank(20),
    p([r('"SECOND PARTY"', {bold:true})]),
    blank(20),
  ];
  buyers.forEach((b, i) => {
    const email = d['buyer'+(i+1)+'_email'] || '';
    const phone = d['buyer'+(i+1)+'_phone'] || '';
    children.push(p([r('________________________________________')]));
    children.push(p([r(b + (email ? '   Email: '+email : '') + (phone ? '   Phone: '+phone : ''))]));
    children.push(p([r('Signature                                      Printed Name')]));
    children.push(blank(20));
  });
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN } }, children }]
  });
}

// ─── Buyer Acknowledgments ────────────────────────────────────────────────────
function buildBuyerAcknowledgments(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('Important Buyer Acknowledgments \u2013 ' + d.property_address),
    blank(40),
    p([r('_____ I understand that my monthly payment of '), r(d.total_monthly_payment || '', {bold:true}), r(' includes principal, interest, estimated taxes, and a '), r(d.servicing_fee || '$120.00', {bold:true}), r(' servicing fee. This amount may change if those costs increase, with proper notice.')]),
    blank(30),
    p([r('_____ I understand that payments are due on the 1st of each month and are considered late after the 5th. A 10% late fee will apply if payment is not received by the 5th.')]),
    blank(30),
    p([r('_____ I understand that if I default, the Seller will give me written notice and I will have 15 days to fix the issue. If I don\u2019t, the Seller may declare the full remaining balance due. If the default continues for 45 days or more, the Seller may pursue forfeiture or foreclosure, and I may be responsible for court costs and administrative fees.')]),
    blank(30),
    p([r('_____ I am purchasing this property '), r('as-is', {bold:true}), r(' and have had the opportunity to conduct my own due diligence or inspections. '), r(d.seller_name, {bold:true}), r(' has made no representations regarding the condition of the property.')]),
    blank(30),
    p([r('_____ I am responsible for maintaining the property, including all required repairs, lawn care, utilities, working smoke detectors at all times, and insurance, with '), r(d.seller_name, {bold:true}), r(' listed as an additional insured.')]),
    blank(30),
    p([r('_____ I understand that '), r(d.seller_name, {bold:true}), r(' is not escrowing taxes or insurance and is not acting in an escrow or trust capacity.')]),
    blank(30),
    p([r('_____ I understand that the seller has done NO inspections and has no knowledge of the condition of the house including, but not limited to plumbing, HVAC, Electrical, Roof, etc., and that any work needed will be '), r('solely', {bold:true}), r(' at the buyers expense.')]),
    blank(30),
    p([r('_____ I agree to assume full responsibility for obtaining a valid certificate of occupancy, if required by local regulations, prior to the property being occupied.')]),
    blank(30),
    p([r('_____ I shall transfer utilities into my name within 48 hours of executing this agreement.')]),
    blank(30),
    p([r('_____ I have received the amortization schedule.')]),
    blank(30),
    p([r('_____ I have received the following required documents:', {bold:true})]),
    p([r('- Seller Disclosure Statement')], {indent:360}),
    p([r('- Lead-Based Paint Disclosure')], {indent:360}),
    p([r('- Radon Disclosure')], {indent:360}),
    p([r('- AS IS Sale Disclosure')], {indent:360}),
    p([r('- Deposit Agreement')], {indent:360}),
    p([r('- Amortization Schedule')], {indent:360}),
    p([r('- Truth In Lending Disclosure')], {indent:360}),
    p([r('- Agreement for Deed')], {indent:360}),
    blank(40),
    hRule(),
  ];
  buyers.forEach((b, i) => {
    const email = d['buyer'+(i+1)+'_email'] || '';
    const phone = d['buyer'+(i+1)+'_phone'] || '';
    children.push(p([r('Buyer Signature: ________________________________________    Date: ________________')]));
    children.push(p([r('Printed Name: '), r(b, {bold:true})]));
    if (email) children.push(p([r('Email: '), r(email, {bold:true}), r('    Phone: '), r(phone, {bold:true})]));
    children.push(blank(30));
  });
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

// ─── TIL — legal size ─────────────────────────────────────────────────────────
function buildTIL(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const loanTermDisplay = d.loan_term_months ? (d.loan_term_months + ' months / ' + Math.round(d.loan_term_months/12) + ' years') : '360 months / 30 years';
  const children = [
    heading('Federal Truth-in-Lending Disclosure Statement', 24),
    blank(30),
    p([r('This disclosure is provided in compliance with the Federal Truth-in-Lending Act. Please review carefully before signing.')]),
    blank(40),
    sectionHead('LOAN INFORMATION:'),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [3400, 6680],
      rows: [
        new TableRow({ children: [cell('Borrower(s) Name(s):', 3400, {bold:true}), cell(buyers.join(', '), 6680)] }),
        new TableRow({ children: [cell('Lender/Seller Name:', 3400, {bold:true}), cell(d.seller_name, 6680)] }),
        new TableRow({ children: [cell('Property Address:', 3400, {bold:true}), cell(d.property_address, 6680)] }),
        new TableRow({ children: [cell('Loan Amount:', 3400, {bold:true}), cell(d.loan_amount_numbers, 6680)] }),
        new TableRow({ children: [cell('Loan Term:', 3400, {bold:true}), cell(loanTermDisplay, 6680)] }),
        new TableRow({ children: [cell('Date of Transaction:', 3400, {bold:true}), cell(d.doc_date, 6680)] }),
      ]
    }),
    blank(40),
    sectionHead('KEY DISCLOSURES:'),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [4600, 5480],
      rows: [
        new TableRow({ children: [cell('Annual Percentage Rate (APR):', 4600, {bold:true}), cell(d.interest_rate_numbers + '%', 5480)] }),
        new TableRow({ children: [cell('Finance Charge (Total cost of credit):', 4600, {bold:true}), cell(d.finance_charge || '', 5480)] }),
        new TableRow({ children: [cell('Amount Financed (Credit provided to borrower):', 4600, {bold:true}), cell(d.loan_amount_numbers, 5480)] }),
        new TableRow({ children: [cell('Total of Payments (Total amount paid after all payments):', 4600, {bold:true}), cell(d.total_of_payments || '', 5480)] }),
      ]
    }),
    blank(40),
    sectionHead('PAYMENT SCHEDULE:'),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [3400, 3340, 3340],
      rows: [
        new TableRow({ children: [cell('Number of Payments', 3400, {bold:true}), cell('Amount of Each Payment', 3340, {bold:true}), cell('When Payments Are Due', 3340, {bold:true})] }),
        new TableRow({ children: [cell(d.loan_term_months || '360', 3400), cell(d.pi_numbers || '', 3340), cell('Monthly, beginning ' + (d.first_payment_date || ''), 3340)] }),
      ]
    }),
    blank(40),
    sectionHead('OTHER TERMS AND CONDITIONS:'),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [3400, 6680],
      rows: [
        new TableRow({ children: [cell('Late Payment Charge:', 3400, {bold:true}), cell('10% of the total monthly payment if not received by the 5th of the month', 6680)] }),
        new TableRow({ children: [cell('Prepayment:', 3400, {bold:true}), cell(CHECK + '  May prepay in full without penalty. No partial prepayments unless mutually agreed.', 6680)] }),
        new TableRow({ children: [cell('Security:', 3400, {bold:true}), cell('This loan is secured by the real property described above', 6680)] }),
        new TableRow({ children: [cell('Assumption:', 3400, {bold:true}), cell('This obligation may not be assumed without prior written consent of the Seller.', 6680)] }),
      ]
    }),
    blank(40),
    sectionHead('ACKNOWLEDGEMENT OF RECEIPT:'),
    p([r('You acknowledge receiving a complete copy of this disclosure before becoming obligated on this loan. You are not required to complete this agreement merely because you have received this disclosure.')]),
    blank(60),
    hRule(),
    blank(20),
  ];
  buyers.forEach(b => {
    children.push(p([r('Borrower Signature: ____________________________    Date: ________________')]));
    children.push(p([r('Printed Name: '), r(b, {bold:true})]));
    children.push(blank(30));
  });
  children.push(p([r('Lender/Seller Signature: ____________________________    Date: ________________')]));
  children.push(p([r(d.seller_name + ' \u2014 Seller/Lender')]));
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN } }, children }]
  });
}

// ─── Buyer Contact Info ───────────────────────────────────────────────────────
function buildBuyerContactInfo(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('Buyer Contact & Emergency Information'),
    blank(40),
    p([r('Property Address: '), r(d.property_address, {bold:true})]),
    blank(30),
  ];
  buyers.forEach((b, i) => {
    const email = d['buyer'+(i+1)+'_email'] || '';
    const phone = d['buyer'+(i+1)+'_phone'] || '';
    children.push(p([r('Buyer '+(i+1)+' Name: '), r(b, {bold:true})]));
    children.push(p([r('Phone: '), r(phone || '_______________________'), r('     Email: '), r(email || '_______________________')]));
    children.push(blank(20));
  });
  children.push(p([r('Mailing Address (if different): _____________________________________________')]));
  children.push(blank(20));
  children.push(p([r('Emergency Contact Name & Phone: _________________________________________________________')]));
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

// ─── Deposit Agreement — legal size ──────────────────────────────────────────
function buildDepositAgreement(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const buyerStr = buyers.join(' and ');
  const children = [
    heading('DEPOSIT AGREEMENT', 24),
    subheading('For Agreement for Deed / Installment Sale'),
    blank(40),
    p([r('This Non-Refundable Deposit Agreement ("Agreement") is made as of '), r(d.doc_date, {bold:true}), r(' between the Seller '), r(d.seller_name, {bold:true}), r(' ("Seller") and Buyer(s) '), r(buyerStr, {bold:true}), r(' ("Buyer").')]),
    blank(30),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [2800, 7280],
      rows: [
        new TableRow({ children: [cell('Property Address:', 2800, {bold:true}), cell(d.property_address, 7280)] }),
        new TableRow({ children: [cell('Deposit Amount:', 2800, {bold:true}), cell(d.deposit_amount_words ? d.deposit_amount_words + ' (' + d.deposit_amount + ')' : (d.deposit_amount || ''), 7280)] }),
        new TableRow({ children: [cell('Closing Deadline:', 2800, {bold:true}), cell(d.closing_date || '', 7280)] }),
      ]
    }),
    blank(30),
    sectionHead('1. Underlying Agreement for Deed'),
    p([r('The parties intend to enter into an Agreement for Deed / Installment Sale Agreement for the property described above. This Deposit Agreement is ancillary to and supplements that Agreement.')]),
    blank(20),
    sectionHead('2. Deposit Terms'),
    p([r('Buyer agrees to pay a non-refundable deposit of '), r(d.deposit_amount_words ? d.deposit_amount_words + ' (' + d.deposit_amount + ')' : (d.deposit_amount || '________________'), {bold:true}), r(' to Seller. Unless otherwise stated in this Agreement, the Deposit is non-refundable and will be applied to the purchase price of the Property at Closing.')]),
    blank(20),
    sectionHead('3. Closing Date'),
    p([r('Buyer agrees to complete the purchase and close under the Agreement for Deed on or before '), r(d.closing_date || '________________', {bold:true}), r('. Time is of the essence with respect to Buyer\u2019s obligation to close on or before the Closing Date.')]),
    blank(20),
    sectionHead('4. Failure to Close / Liquidated Damages'),
    p([r('If Buyer fails to close on or before the Closing Date, and such failure is not caused by Seller\u2019s default or inability to perform: (a) Buyer\u2019s Deposit shall be fully earned by Seller and non-refundable; (b) the Deposit shall serve as liquidated damages, not as a penalty; (c) upon Seller\u2019s retention of the Deposit, Seller shall be released from any further obligation to sell the Property to Buyer; and (d) Seller may market or sell the Property to another buyer.')]),
    blank(20),
    sectionHead('5. Representations'),
    p([r('Buyer represents that Buyer has had the opportunity to inspect the Property and is satisfied with the condition of the Property, or is willing to accept it "as-is, where-is", subject to the terms of the Agreement for Deed.')]),
    blank(20),
    sectionHead('6. Governing Law'),
    p([r('This Agreement shall be governed by and construed in accordance with the laws of the State of '), r(d.state, {bold:true}), r('.')]),
    blank(20),
    sectionHead('7. Entire Agreement'),
    p([r('This Agreement contains the entire understanding between the parties regarding the Deposit and Closing Date and supplements (but does not replace) the Agreement for Deed. In the event of a conflict, the terms of this Agreement shall control regarding the Deposit, unless otherwise agreed in writing.')]),
    blank(20),
    sectionHead('8. Electronic Signatures'),
    p([r('Signatures transmitted by electronic means (including scanned or digitally signed copies) shall be deemed original signatures and fully binding on the parties.')]),
    blank(50),
    hRule(),
    blank(20),
    p([r(d.seller_name + ' \u2014 Seller', {bold:true})]),
    p([r('Signature: ____________________________    Date: ________________')]),
    blank(30),
  ];
  buyers.forEach(b => {
    children.push(p([r(b + ' \u2014 Buyer', {bold:true})]));
    children.push(p([r('Signature: ____________________________    Date: ________________')]));
    children.push(blank(30));
  });
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN } }, children }]
  });
}

// ─── AS-IS Addendum — one page legal ─────────────────────────────────────────
function buildAsIsAddendum(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('AS-IS CONDITION & DUE DILIGENCE ADDENDUM', 22),
    blank(20),
    p([r('This As-Is Condition and Due Diligence Addendum ("Addendum") is incorporated into and made part of the Agreement for Deed between '), r(buyers.join(' and '), {bold:true}), r(' ("Buyer") and '), r(d.seller_name, {bold:true}), r(' ("Seller") for the property located at '), r(d.property_address, {bold:true}), r(', dated '), r(d.doc_date, {bold:true}), r('.')]),
    blank(20),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [FULL],
      rows: [
        new TableRow({ children: [new TableCell({ borders: bdr(), width: {size: FULL, type: WidthType.DXA}, margins: {top:80, bottom:80, left:120, right:120}, children: [
          p([r('1. AS-IS SALE', {bold:true, underline:true})], {after:40}),
          p([r('Buyer acknowledges the Property is being sold strictly AS-IS, WHERE-IS with all faults, whether known or unknown. Seller makes no warranties or representations, express or implied, regarding condition, habitability, value, or future performance.')], {after:60}),
          p([r('2. SELLER NON-OCCUPANCY & LIMITED KNOWLEDGE', {bold:true, underline:true})], {after:40}),
          p([r('Buyer acknowledges Seller is an investor who has never occupied the Property and may have never personally visited it. Seller has limited or no first-hand knowledge of its condition.')], {after:60}),
          p([r('3. NO RELIANCE', {bold:true, underline:true})], {after:40}),
          p([r('Buyer confirms they are not relying on statements, marketing materials, or opinions from Seller or agents. Buyer relies solely on independent investigations.')], {after:60}),
          p([r('4. BUYER DUE DILIGENCE', {bold:true, underline:true})], {after:40}),
          p([r('Buyer is strongly encouraged to obtain inspections including general, structural, pest, environmental, and contractor evaluations. All inspections are Buyer\u2019s sole responsibility and expense.')], {after:60}),
          p([r('5. ACCEPTANCE OF CONDITION', {bold:true, underline:true})], {after:40}),
          p([r('Buyer affirms adequate opportunity to inspect and fully accepts the Property condition. All repairs and improvements are Buyer\u2019s responsibility.')], {after:60}),
          p([r('6. RELEASE', {bold:true, underline:true})], {after:40}),
          p([r('Buyer releases Seller from any and all claims related to condition, defects, code violations, or repairs discovered after closing or possession.')], {after:60}),
          p([r('7. VOLUNTARY AGREEMENT', {bold:true, underline:true})], {after:40}),
          p([r('Buyer confirms this Addendum is signed freely, without duress, and is a material part of the purchase agreement.')], {after:0}),
        ]})]}),
      ]
    }),
    blank(20),
    hRule(),
    blank(20),
  ];
  buyers.forEach((b, i) => {
    const email = d['buyer'+(i+1)+'_email'] || '';
    const phone = d['buyer'+(i+1)+'_phone'] || '';
    children.push(p([r('Buyer Initials: ________    Date: ________________    Buyer Signature: ____________________________    Date: ________________')]));
    children.push(p([r(b + (email ? '   |   '+email : '') + (phone ? '   |   '+phone : ''), {bold:true})]));
    children.push(blank(20));
  });
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN_TIGHT } }, children }]
  });
}

// ─── Lead Paint Disclosure — one page legal ───────────────────────────────────
function buildLeadPaint(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('Disclosure of Information on Lead-Based Paint and/or Lead-Based Paint Hazards', 21),
    blank(20),
    p([r('Lead Warning Statement: ', {bold:true}), r('Every purchaser of any interest in residential real property on which a residential dwelling was built prior to 1978 is notified that such property may present exposure to lead from lead-based paint that may place young children at risk of developing lead poisoning. Lead poisoning in young children may produce permanent neurological damage, including learning disabilities, reduced intelligence quotient, behavioral problems, and impaired memory. Lead poisoning also poses a particular risk to pregnant women. The seller is required to provide the buyer with any information on lead-based paint hazards and notify the buyer of any known hazards. A risk assessment or inspection is recommended prior to purchase.')], {line: 264}),
    blank(20),
    p([r('Property Address: '), r(d.property_address, {bold:true, underline:true})]),
    blank(20),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [FULL],
      rows: [
        new TableRow({ children: [new TableCell({ borders: bdr(), width: {size: FULL, type: WidthType.DXA}, margins: {top:80, bottom:80, left:120, right:120}, children: [
          p([r('SELLER\'S DISCLOSURE (Initial)', {bold:true})], {after:30}),
          p([r('_____  (a) Presence of lead-based paint: '), r('[  ] Known hazards are present (explain below).   ['+CHECK+'] Seller has no knowledge of lead-based paint hazards in the housing.')], {after:20}),
          p([r('_____  (b) Records and reports: '), r('[  ] Seller has provided all available records.   ['+CHECK+'] Seller has no reports or records pertaining to lead-based paint hazards in the housing.')], {after:60}),
          p([r('PURCHASER\'S ACKNOWLEDGMENT (Initial)', {bold:true})], {after:30}),
          p([r('_N/A_  (c) Purchaser has received copies of all information listed above.')], {after:20}),
          p([r('___/___  (d) Purchaser has received the pamphlet '), r('Protect Your Family from Lead in Your Home', {italics:true}), r('.')], {after:20}),
          p([r('___/___  (e) Purchaser has: '), r('[  ] received a 10-day opportunity to conduct a risk assessment or inspection; or   ['+CHECK+'] waived the opportunity to conduct a risk assessment or inspection.')], {after:60}),
          p([r('AGENT\'S ACKNOWLEDGMENT (Initial)', {bold:true})], {after:30}),
          p([r('_N/A_  (f) Agent has informed the seller of the seller\u2019s obligations under 42 U.S.C. 4852d and is aware of his/her responsibility to ensure compliance.')], {after:0}),
        ]})]}),
      ]
    }),
    blank(20),
    p([r('Certification of Accuracy: ', {bold:true}), r('The following parties have reviewed the information above and certify, to the best of their knowledge, that the information they have provided is true and accurate.')]),
    blank(20),
    hRule(),
    blank(20),
    p([r('Seller: ____________________________    Date: ________________    ('), r(d.seller_name, {bold:true}), r(')')]),
    blank(20),
  ];
  buyers.forEach(b => {
    children.push(p([r('Buyer: ____________________________    Date: ________________    ('), r(b, {bold:true}), r(')')]));
    children.push(blank(20));
  });
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN_TIGHT } }, children }]
  });
}

// ─── Radon Disclosure — one page legal ───────────────────────────────────────
function buildRadonDisclosure(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('DISCLOSURE OF INFORMATION ON RADON HAZARDS', 22),
    subheading('For Current and Prospective Buyers'),
    blank(20),
    p([r('Radon Warning Statement: ', {bold:true}), r('Each buyer in this residence is notified that the property may present exposure to levels of indoor radon gas that may place occupants at risk of developing radon-induced lung cancer. Radon is a Class-A human carcinogen and the leading cause of lung cancer in nonsmokers. The seller is required to provide each buyer with any information on radon test results that present a radon hazard. It is strongly recommended that ALL properties have a radon test performed and radon hazards mitigated if elevated levels are found.')], {line: 264}),
    blank(20),
    p([r('Dwelling Unit Address: '), r(d.property_address, {bold:true})]),
    blank(20),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [FULL],
      rows: [
        new TableRow({ children: [new TableCell({ borders: bdr(), width: {size: FULL, type: WidthType.DXA}, margins: {top:80, bottom:80, left:120, right:120}, children: [
          p([r('SELLER\'S DISCLOSURE (initial each that applies)', {bold:true})], {after:30}),
          p([r('_____  (a) Seller has no knowledge of elevated radon concentrations (or records or reports pertaining to elevated radon concentrations) in the dwelling unit.')], {after:20}),
          p([r('_N/A_  (b) Radon concentrations at or above the recommended Radon Action Level of 4.0 pCi/L are known to be present within the dwelling unit.')], {after:20}),
          p([r('_____  (c) Seller has provided the Buyer with copies of all available records and reports, if any, pertaining to radon concentrations within the dwelling unit.')], {after:60}),
          p([r('BUYER\'S ACKNOWLEDGMENT (initial each that applies)', {bold:true})], {after:30}),
          p([r('___/___  (d) Buyer has received copies of all information listed above.')], {after:20}),
          p([r('___/___  (e) Buyer has received the pamphlet "Radon Guide for Tenants".')], {after:60}),
          p([r('AGENT\'S ACKNOWLEDGMENT (initial if applicable)', {bold:true})], {after:30}),
          p([r('_N/A_  (g) Agent has informed the seller of the seller\u2019s obligations under applicable state law.')], {after:0}),
        ]})]}),
      ]
    }),
    blank(20),
    p([r('Certification of Accuracy: ', {bold:true}), r('The following parties have reviewed the information above and each party certifies, to the best of his or her knowledge, that the information he or she provided is true and accurate.')]),
    blank(20),
    hRule(),
    blank(20),
    p([r('Seller: ____________________________    Date: ________________    ('), r(d.seller_name, {bold:true}), r(')')]),
    blank(20),
  ];
  buyers.forEach(b => {
    children.push(p([r('Buyer: ____________________________    Date: ________________    ('), r(b, {bold:true}), r(')')]));
    children.push(blank(20));
  });
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN_TIGHT } }, children }]
  });
}

// ─── Memorandum — one page legal, full notary blocks ─────────────────────────
function buildMemorandum(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const buyerStr = buyers.join(' and ');
  const yr = new Date().getFullYear();
  const children = [
    heading('MEMORANDUM OF AGREEMENT FOR DEED', 22),
    blank(20),
    p([r('THIS IS TO CERTIFY that an Agreement for Deed between '), r(d.seller_name, {bold:true}), r(', designated "SELLER", and '), r(buyerStr, {bold:true}), r(', designated "BUYER", was entered into on '), r(d.doc_date, {bold:true}), r(', wherein the Seller agreed to sell and the Buyer agreed to purchase the following described real estate:')]),
    blank(20),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [2400, 7680],
      rows: [
        new TableRow({ children: [cell('Property Address:', 2400, {bold:true}), cell(d.property_address, 7680)] }),
        new TableRow({ children: [cell('Legal Description:', 2400, {bold:true}), cell(d.legal_description || '', 7680)] }),
        new TableRow({ children: [cell('P.I.N. #:', 2400, {bold:true}), cell(d.pid || '', 7680)] }),
        new TableRow({ children: [cell('Purchase Price:', 2400, {bold:true}), cell(d.purchase_price_words ? d.purchase_price_words + ' (' + d.purchase_price_numbers + ')' : (d.purchase_price_numbers || ''), 7680)] }),
      ]
    }),
    blank(20),
    p([r('Subject to easements, restrictions, and reservations of former instruments of record, if any. Subject to general taxes for '), r(yr.toString(), {bold:true}), r(' payable in '), r((yr+1).toString(), {bold:true}), r(' and thereafter.')]),
    blank(20),
    p([r('The undersigned further certify that said Agreement contains an acknowledgment by Buyer of Buyer\u2019s Special Warranty Deed reconveying the above described real estate to Seller, authorizing reliance by third parties upon said deed, if recorded, as conclusive evidence of the cancellation of said Agreement and reconveyance of all of Buyer\u2019s then interest in said real estate to Seller.')]),
    blank(30),
    hRule(),
    blank(20),
    // Seller block
    p([r('SELLER: ', {bold:true}), r(d.seller_name)]),
    blank(10),
    p([r('By: ____________________________    Title: Member/Manager    Date: ________________')]),
    p([r('Printed Name: '), r(d.signor || '', {bold:true})]),
    blank(30),
    // Seller notary
    p([r('STATE OF '), r(d.state || '______________', {bold:true}), r('       )'), ]),
    p([r('COUNTY OF '), r(d.county || '______________', {bold:true}), r('     )    SS.')]),
    blank(10),
    p([r('On this ______ day of _______________, '), r(yr.toString()), r(', before me, the undersigned Notary Public, personally appeared '), r(d.signor || '__________________________', {bold:true}), r(', known to me (or satisfactorily proven) to be the Member/Manager of '), r(d.seller_name, {bold:true}), r(', and acknowledged that he/she executed the foregoing instrument on behalf of said entity for the purposes therein contained.')]),
    blank(10),
    p([r('Notary Public: ____________________________    Commission Expires: ________________')]),
    p([r('My Commission is for the State of: ____________________________    (SEAL)')]),
    blank(30),
    hRule(),
    blank(20),
  ];

  // Buyer blocks
  buyers.forEach((b, i) => {
    const email = d['buyer'+(i+1)+'_email'] || '';
    const phone = d['buyer'+(i+1)+'_phone'] || '';
    children.push(p([r('BUYER: ', {bold:true}), r(b)]));
    if (email || phone) children.push(p([r((email ? 'Email: '+email : '') + (phone ? '   Phone: '+phone : ''))]));
    children.push(blank(10));
    children.push(p([r('Signature: ____________________________    Date: ________________')]));
    children.push(p([r('Printed Name: '), r(b, {bold:true})]));
    children.push(blank(20));
    children.push(p([r('STATE OF '), r(d.state || '______________', {bold:true}), r('       )')]));
    children.push(p([r('COUNTY OF '), r(d.county || '______________', {bold:true}), r('     )    SS.')]));
    children.push(blank(10));
    children.push(p([r('On this ______ day of _______________, '), r(yr.toString()), r(', before me, the undersigned Notary Public, personally appeared '), r(b, {bold:true}), r(', known to me (or satisfactorily proven) to be the person whose name is subscribed to the within instrument and acknowledged that he/she executed the same for the purposes therein contained.')]));
    children.push(blank(10));
    children.push(p([r('Notary Public: ____________________________    Commission Expires: ________________')]));
    children.push(p([r('My Commission is for the State of: ____________________________    (SEAL)')]));
    children.push(blank(20));
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN_TIGHT } }, children }]
  });
}

// ─── Seller Disclosure — 2 pages legal, checkmarks in NO column ──────────────
function buildSellerDisclosure(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const items = [
    'Seller has occupied the property within the last 12 months. (If "no," please identify capacity or explain the relationship to property.)',
    'I currently have flood hazard insurance on the property.',
    'I am aware of flooding or recurring leakage problems in the crawl space or basement.',
    'I am aware that the property is located in a floodplain.',
    'I am aware of material defects in the basement or foundation (including cracks and bulges).',
    'I am aware of leaks or material defects in the roof, ceilings, or chimney.',
    'I am aware of material defects in the walls, windows, doors, or floors.',
    'I am aware of material defects in the electrical system.',
    'I am aware of material defects in the plumbing system (includes water heater, sump pump, water treatment system, sprinkler system, and swimming pool).',
    'I am aware of material defects in the well or well equipment.',
    'I am aware of unsafe conditions in the drinking water.',
    'I am aware of material defects in the heating, air conditioning, or ventilating systems.',
    'I am aware of material defects in the fireplace or wood burning stove.',
    'I am aware of material defects in the septic, sanitary sewer, or other disposal system.',
    'I am aware of unsafe concentrations of radon on the premises.',
    'I am aware of unsafe concentrations of or unsafe conditions relating to asbestos on the premises.',
    'I am aware of unsafe concentrations of or unsafe conditions relating to lead paint, lead water pipes, lead plumbing pipes or lead in the soil on the premises.',
    'I am aware of mine subsidence, underground pits, settlement, sliding, upheaval, or other earth stability defects on the premises.',
    'I am aware of current infestations of termites or other wood boring insects.',
    'I am aware of a structural defect caused by previous infestations of termites or other wood boring insects.',
    'I am aware of underground fuel storage tanks on the property.',
    'I am aware of boundary or lot line disputes.',
    'I have received notice of violation of local, state or federal laws or regulations relating to this property, which violation has not been corrected.',
    'I am aware that this property has been used for the manufacture of methamphetamine as defined in Section 10 of the Methamphetamine Control and Community Protection Act.',
  ];

  const headerRow = new TableRow({ children: [
    cell('YES', 720, {bold:true, align: AlignmentType.CENTER, size:18}),
    cell('NO',  720, {bold:true, align: AlignmentType.CENTER, size:18}),
    cell('N/A', 720, {bold:true, align: AlignmentType.CENTER, size:18}),
    cell('DISCLOSURE ITEM', 7920, {bold:true, size:18}),
  ]});

  const dataRows = items.map((item, i) =>
    new TableRow({ children: [
      cell('', 720, {align: AlignmentType.CENTER}),
      cell(CHECK, 720, {align: AlignmentType.CENTER, size:16}),
      cell('', 720, {align: AlignmentType.CENTER}),
      cell((i+1) + '.  ' + item, 7920, {size:17}),
    ]})
  );

  const children = [
    heading('RESIDENTIAL REAL PROPERTY DISCLOSURE REPORT', 21),
    blank(10),
    p([r('NOTICE: THE PURPOSE OF THIS REPORT IS TO PROVIDE PROSPECTIVE BUYERS WITH INFORMATION ABOUT MATERIAL DEFECTS IN THE RESIDENTIAL REAL PROPERTY BEFORE THE SIGNING OF A CONTRACT. THIS REPORT DOES NOT LIMIT THE PARTIES\u2019 RIGHT TO CONTRACT FOR THE SALE OF RESIDENTIAL REAL PROPERTY IN "AS IS" CONDITION. COMPLETION OF THIS REPORT BY THE SELLER CREATES LEGAL OBLIGATIONS; THEREFORE THE SELLER MAY WISH TO CONSULT AN ATTORNEY PRIOR TO COMPLETION.'), ], {align: AlignmentType.CENTER, after: 60, size: 18}),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [2400, 7680],
      rows: [
        new TableRow({ children: [cell('Property Address:', 2400, {bold:true, size:18}), cell(d.property_address, 7680, {size:18})] }),
        new TableRow({ children: [cell('Seller\'s Name:', 2400, {bold:true, size:18}), cell(d.seller_name, 7680, {size:18})] }),
        new TableRow({ children: [cell('Date of Report:', 2400, {bold:true, size:18}), cell(d.doc_date, 7680, {size:18})] }),
      ]
    }),
    blank(20),
    p([r('The seller represents that to the best of his or her actual knowledge, the following statements have been accurately noted as "YES" (correct), "NO" (incorrect), or "N/A" (not applicable). If the response to any statement is "YES" or "N/A", the seller shall provide explanation below.', {size:18})]),
    blank(10),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [720, 720, 720, 7920],
      rows: [headerRow, ...dataRows]
    }),
    blank(20),
    p([r('Explanation / Additional Information (for any YES or N/A responses):', {bold:true, size:18})]),
    p([r('_________________________________________________________________________')]),
    p([r('_________________________________________________________________________')]),
    blank(20),
    p([r('Seller certifies that the information provided is based on actual notice or actual knowledge without specific investigation. ', {size:18}), r('THE SELLER ACKNOWLEDGES A CONTINUING OBLIGATION TO SUPPLEMENT THIS DISCLOSURE PRIOR TO CLOSING.', {bold:true, size:18})]),
    blank(20),
    p([r('Seller: ____________________________    Date: ________________    ('), r(d.seller_name, {bold:true}), r(')')]),
    blank(20),
    p([r('THE PROSPECTIVE BUYER IS AWARE THAT THE PARTIES MAY CHOOSE TO NEGOTIATE A SALE "AS IS." THIS DISCLOSURE IS NOT A SUBSTITUTE FOR INSPECTIONS. THE FACT THAT THE SELLER IS NOT AWARE OF A PARTICULAR CONDITION IS NO GUARANTEE THAT IT DOES NOT EXIST.', {bold:true, size:18})]),
    blank(20),
    hRule(),
    blank(10),
  ];
  buyers.forEach(b => {
    children.push(p([r('Buyer: ____________________________    Date: ________________    ('), r(b, {bold:true}), r(')')]));
    children.push(blank(20));
  });
  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 19 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN_TIGHT } }, children }]
  });
}

// ─── Main export ──────────────────────────────────────────────────────────────
module.exports = async function generateDocs(data) {
  const docs = [
    { name: '01_Agreement_for_Deed.docx',      doc: buildAgreementForDeed(data) },
    { name: '02_Buyer_Acknowledgments.docx',   doc: buildBuyerAcknowledgments(data) },
    { name: '03_Truth_in_Lending.docx',        doc: buildTIL(data) },
    { name: '04_Buyer_Contact_Info.docx',      doc: buildBuyerContactInfo(data) },
    { name: '05_Deposit_Agreement.docx',       doc: buildDepositAgreement(data) },
    { name: '06_AS_IS_Addendum.docx',          doc: buildAsIsAddendum(data) },
    { name: '07_Lead_Paint_Disclosure.docx',   doc: buildLeadPaint(data) },
    { name: '08_Radon_Disclosure.docx',        doc: buildRadonDisclosure(data) },
    { name: '09_Memorandum_of_Agreement.docx', doc: buildMemorandum(data) },
    { name: '10_Seller_Disclosure.docx',       doc: buildSellerDisclosure(data) },
    { name: '11_Buyer_Acknowledgments.docx',   doc: buildBuyerAcknowledgments(data) },
  ];
  // Remove duplicate - fix list
  const finalDocs = [
    { name: '01_Agreement_for_Deed.docx',      doc: buildAgreementForDeed(data) },
    { name: '02_Buyer_Acknowledgments.docx',   doc: buildBuyerAcknowledgments(data) },
    { name: '03_Truth_in_Lending.docx',        doc: buildTIL(data) },
    { name: '04_Buyer_Contact_Info.docx',      doc: buildBuyerContactInfo(data) },
    { name: '05_Deposit_Agreement.docx',       doc: buildDepositAgreement(data) },
    { name: '06_AS_IS_Addendum.docx',          doc: buildAsIsAddendum(data) },
    { name: '07_Lead_Paint_Disclosure.docx',   doc: buildLeadPaint(data) },
    { name: '08_Radon_Disclosure.docx',        doc: buildRadonDisclosure(data) },
    { name: '09_Memorandum_of_Agreement.docx', doc: buildMemorandum(data) },
    { name: '10_Seller_Disclosure.docx',       doc: buildSellerDisclosure(data) },
  ];
  const buffers = await Promise.all(finalDocs.map(({ doc }) => Packer.toBuffer(doc)));
  return new Promise((resolve, reject) => {
    const archive = archiver('zip', { zlib: { level: 6 } });
    const chunks = [];
    archive.on('data', chunk => chunks.push(chunk));
    archive.on('end', () => resolve(Buffer.concat(chunks)));
    archive.on('error', reject);
    finalDocs.forEach(({ name }, i) => archive.append(buffers[i], { name }));
    archive.finalize();
  });
};
