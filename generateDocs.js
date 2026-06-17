const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, HeadingLevel, PageBreak,
  TabStopType, TabStopPosition, UnderlineType
} = require('docx');
const archiver = require('archiver');
const { Readable } = require('stream');

// ─── helpers ──────────────────────────────────────────────────────────────────

const ARIAL = 'Arial';
const LETTER = { width: 12240, height: 15840 };
const LEGAL  = { width: 12240, height: 20160 };
const MARGIN = { top: 1080, right: 1080, bottom: 1080, left: 1080 }; // 0.75"
const FULL = 10080; // content width with 0.75" margins on letter

function r(text, opts = {}) {
  return new TextRun({ text, font: ARIAL, size: opts.size || 20, bold: opts.bold || false,
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
    italics: opts.italics || false });
}
function p(children, opts = {}) {
  if (typeof children === 'string') children = [r(children)];
  return new Paragraph({
    children,
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after || 120, line: opts.line || 276 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    tabStops: opts.tabStops || undefined,
    pageBreakBefore: opts.pageBreak || false,
    border: opts.border || undefined,
  });
}
function blank(before, after) { return p('', { before: before || 0, after: after || 60 }); }
function sigLine(label, width) {
  const w = width || 3600;
  return new TableCell({
    borders: noBorder(),
    width: { size: w, type: WidthType.DXA },
    children: [
      p([r('________________________________________', { size: 20 })]),
      p([r(label, { size: 18 })]),
    ]
  });
}
function noBorder() {
  const n = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top: n, bottom: n, left: n, right: n };
}
function sigTable(cells) {
  return new Table({
    width: { size: FULL, type: WidthType.DXA },
    columnWidths: cells.map(c => c._cellProperties ? c._cellProperties.width.size : 3360),
    rows: [new TableRow({ children: cells })],
  });
}
function hRule() {
  return p('', { border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } }, before: 60, after: 60 });
}
function heading(text) {
  return p([r(text, { bold: true, size: 22 })], { align: AlignmentType.CENTER, before: 120, after: 80 });
}
function sectionHead(text) {
  return p([r(text, { bold: true, size: 20, underline: true })], { before: 100, after: 60 });
}

// ─── document builders ────────────────────────────────────────────────────────

function buildAgreementForDeed(d) {
  // Build buyer list
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const buyerStr = buyers.join(' and ');
  const buyerLines = buyers.map((b, i) => {
    const email = d['buyer' + (i+1) + '_email'] || '';
    const phone = d['buyer' + (i+1) + '_phone'] || '';
    return p([r(b + (email ? '  |  ' + email : '') + (phone ? '  |  ' + phone : ''), { size: 19 })]);
  });

  const children = [
    heading('AGREEMENT FOR DEED'),
    blank(),
    p([r('THIS AGREEMENT FOR DEED is made and entered into this '), r(d.doc_date, { bold: true }), r(', by and between, '), r(d.seller_name, { bold: true }), r(' (hereinafter referred to as "First Party"), and '), r(buyerStr, { bold: true }), r(' (hereinafter referred to as "Second Party").')]),
    blank(),
    p([r('WITNESSETH, that if the Second Party shall first make all of the payments and fully perform the covenants hereinafter mentioned on his part to be made and performed, the First Party hereby covenants and agrees to convey to the Second Party, his heirs, executors, administrators, personal representatives, or assigns, in fee simple absolute, clear of all encumbrances, except as set out herein, by a good and sufficient warranty deed, that parcel of land situated in the City of '), r(d.city, { bold: true }), r(', State of '), r(d.state, { bold: true }), r(', known as:')]),
    blank(),
    p([r(d.property_address, { bold: true })], { align: AlignmentType.CENTER }),
    p([r(d.legal_description || '')], { align: AlignmentType.CENTER }),
    p([r('(the "Property")')], { align: AlignmentType.CENTER }),
    blank(),
    p([r('1.', { bold: true }), r('  The Second Party herein covenants and agrees to pay to the First Party the sum of '), r(d.purchase_price_words, { bold: true }), r(' ('), r(d.purchase_price_numbers, { bold: true }), r('), in the following manner:')]),
    blank(),
    p([r('a.  '), r(d.down_payment_words, { bold: true }), r(' ('), r(d.down_payment_numbers, { bold: true }), r(') paid to the First Party. Down payment is to be made out to '), r(d.seller_name, { bold: true }), r('.')], { indent: 360 }),
    blank(),
    p([r('b.  The principal sum of '), r(d.loan_amount_words, { bold: true }), r(' ('), r(d.loan_amount_numbers, { bold: true }), r('), with interest thereon at the rate of '), r(d.interest_rate_numbers, { bold: true }), r('% percent per annum, in monthly installments of '), r(d.pi_words, { bold: true }), r(' ('), r(d.pi_numbers, { bold: true }), r('), beginning on the '), r(d.first_payment_date, { bold: true }), r(', and continuing on the 1st day of each and every month thereafter, until Paid in full or until all principal has been paid in full or until termination of this agreement, as provided hereinafter, whichever shall occur first. Payments are considered late after the 5th of the month and will incur a 10% late fee. ALL Monthly Payments are to be made out to '), r(d.seller_name, { bold: true }), r(' and payable by: '), r(d.payment_method || 'Buildium', { bold: true }), r('. This Agreement is amortized from '), r(d.first_payment_date, { bold: true }), r(' to '), r(d.maturity_date, { bold: true }), r(', with final payment due on '), r(d.maturity_date, { bold: true }), r(', unless prepaid beforehand. Property Taxes for '), r(new Date().getFullYear().toString(), { bold: true }), r(' are '), r(d.monthly_taxes_numbers, { bold: true }), r(' per month. Property Taxes shall be reviewed yearly and the costs amended accordingly each year and paid by the second party at cost. Loan servicing fee of '), r(d.servicing_fee || '$120.00', { bold: true }), r(' charged each month.')], { indent: 360 }),
    blank(),
    p([r('2.  Conveyance of the Property shall be by a special warranty deed upon completion of all payments and full performance by the Second Party as required herein. The second party will not receive any tax documents from the first party and shall rely solely on the amortization schedule for tax purposes. The First Party grants the right of possession and occupancy of the Property to the Second Party upon acceptance of this agreement.')]),
    blank(),
    p([r('3.  Upon completion of all payments and full performance by the Second Party as required herein, title shall be conveyed free and clear of all encumbrances except, any easements, restrictions, limitations, reservations, covenants and conditions of record not coupled with a possibility of reverter, right of reentry or other reverter right which amounts to a qualification of the fee, and subject also to applicable zoning ordinances and real estate taxes for the year in which the deed is delivered, and thereafter.')]),
    blank(),
    p([r('4.  The Second Party agrees to pay Full Replacement Insurance on the property and is required to contact an agent within 72 hours to obtain a homeowners policy. The Second Party shall obtain an insurance policy meeting the requirements set forth herein on or before 10 days from ratification of this agreement. The First Party shall be listed on the policy as an additional insured party. A default of insurance shall be considered a default of this Agreement. The First Party shall remain listed as an additional insured party until completion of all payments by Second Party and delivery of the title by First Party.')]),
    blank(),
    p([r('5.  It is understood that this is a business transaction that carries a level of risk and that the Second Party is not purchasing the property but paying the First Party for possession and use of the property, as well as the potential future right to become the owner of the Property.')]),
    blank(),
    p([r('6.  In the event that the property is taken by eminent domain during the existence of this contract, the First Party will be paid the remaining principal balance due at that time and 100% of the remaining proceeds will be paid the Second Party.')]),
    blank(),
    p([r('7.  The Second Party may prepay the FULL principal balance outstanding at any time without penalty and without notice. Such prepayment shall not include unearned interest. Upon full prepayment, the First Party shall have twenty (20) days in which to deliver a warranty deed. NO PARTIAL PRE-PAYMENTS UNLESS MUTUALLY AGREED UPON.')]),
    blank(),
    p([r('8.  In case of the failure of the Second Party to make payments or any part thereof, or to perform any of the covenants on his part hereby made and entered into, this Agreement shall, at the option of the First Party, be forfeited and terminated, and the Second Party shall forfeit all payments made by him to date on this agreement and said amount shall be retained by the First Party in full satisfaction and liquidation of all damages sustained by the First Party. At such time, this Agreement shall be terminated and rendered null and void. The First Party shall have the right to demand immediate possession of the Property and commence eviction proceedings, if necessary. Any additional payments made by the Second Party to the First Party shall be considered rent and the Second Party may be considered a month-to-month tenant, in the sole and absolute discretion of the First Party. All costs and expenses incurred by the First Party, including but not limited to reasonable attorneys\u2019 fees, shall be paid by the Second Party to the First Party.')]),
    blank(),
    p([r('9.  The time of each payment shall be an essential part of this contract. Any payment not received within ten (10) days of its due date or the accumulation of any three payments not received within five (5) days of their respective due dates, shall constitute a default under paragraph 8 hereinabove.')]),
    blank(),
    p([r('10. Any notice necessary under this Agreement may be sent by mail to the last known address of the party to be notified.')]),
    blank(),
    p([r('11. All covenants and agreements herein contained shall extend to and be obligatory upon the successors, heirs, executors, administrators, personal representatives and assigns of the respective parties.')]),
    blank(),
    p([r('12. The Second Party will not permit, commit or suffer waste to the Property and will maintain the Property and all improvements of the Property at all times in a state of good repair and condition, and will not do or permit to be done anything to the Property that will in any way impair or weaken the value of the Property. In case of the refusal, neglect or inability of the Second Party to repair and maintain said Property, the First Party may, at the First Party\u2019s option, make such repairs or cause the same to be made, and advance money in that behalf, which sums advanced or costs of repairs shall be the obligation of the Second Party and shall be secured by this Agreement. Additionally, the refusal, neglect or inability of the Second Party to repair and maintain said Property shall be considered and event of default and the First Party may exercise all rights set forth in Paragraph 8 above.')]),
    blank(),
    p([r('13. The words "First Party" and "Second Party" herein employed shall be construed to include the plural as well as the singular, and the masculine shall include the feminine and neuter where the context so admits or requires.')]),
    blank(),
    p([r('14. This agreement, and all transactions contemplated hereby, shall be governed by, construed and enforced in accordance with the laws of the State of '), r(d.state, { bold: true }), r('. The parties herein waive trial by jury and agree to submit to the personal jurisdiction and venue of a court of subject matter jurisdiction located in '), r(d.county, { bold: true }), r(', State of '), r(d.state, { bold: true }), r('. In the event that litigation results from or arises out of this Agreement or the performance thereof, the parties agree to reimburse the prevailing party\u2019s reasonable attorney\u2019s fees, court costs, and all other expenses, whether or not taxable by the court as costs, in addition to any other relief to which the prevailing party may be entitled.')]),
    blank(),
    p([r('15. The First Party may sign and record a Notice of Termination of said agreement if the Second Party defaults in the performance of the Second Party\u2019s obligations and responsibilities under this Agreement and such termination shall be effective if no legal proceeding is instituted and no lis pendens is filed by the Second Party within three months of the date such Notice of Termination was recorded.')]),
    blank(),
    p([r('16. Unless specifically disallowed by law, should litigation arise hereunder, service of process therefor may be obtained through certified mail, return receipt requested; the parties hereto waiving any and all rights they may have to object to the method by which service was perfected.')]),
    blank(),
    p([r('17. The First Party may inspect the property with a 72 hour verbal notice every 6 months to determine the condition of the Property. The Second Party agrees to maintain the current level of condition, or better, of the Property. If an inspection shows any damage or degradation of the Property the First Party can give a 30 day notice for the Second Party to repair or replace, as the First Party reasonably deems necessary. If the Second Party fails to promptly make such repair or replacement and has not returned the Property to its original or better condition, then the Second Party shall be considered to be in default. In the event of such default, the First Party may exercise all rights set forth in Paragraph 8 above.')]),
    blank(),
    p([r('18. Notification of Serious Building Problems: The Second Party agrees to notify the First Party immediately upon first discovering any signs of serious building problems such as a crack in the foundation, a tilting porch, a crack in the plaster or stucco, moisture in the ceiling, buckling sheetrock or siding, leaky roof, a spongy floor, a leaky water heater or termite activity.')]),
    blank(),
    p([r('19. Lawful Use: The Second Party agrees that they will not themselves engage in any illegal activities on the premises nor will they allow others to engage in any illegal activities on the premises insofar as they have the power to stop such activities.')]),
    blank(),
    p([r('20. Insurance Consideration: The Second Party agrees that they will keep the property insured for full replacement cost and that they will keep First Party listed as additional insured until the property has been paid in full.')]),
    blank(),
    p([r('21. Alarm System: The Second party agrees to assume over the Alarm System kept by First Party and must maintain the monitoring services for the duration of this agreement.')]),
    blank(),
    p([r('22. The First Party shall not be responsible for any repairs. The Second Party shall be responsible for any and all repairs and maintenance to the Property at their sole cost and expense.')]),
    blank(),
    p([r('23. The Second Party agrees to transfer all utilities to their names within 48 hours after move-in date.')]),
    blank(),
    p([r('24. Acknowledgment: The Second Party agrees hereby acknowledges that they have read this Agreement, understand it, agree to it and have been given a copy. They further have been advised to seek legal, tax, technical expertise and any other counsel of their choosing concerning this contract prior to signing.')]),
    blank(),
    hRule(),
    p([r('DISCLAIMER:', { bold: true })]),
    p([r('By signing below, you hereby acknowledge that this transaction is not a traditional sale and purchase of real property. You (the "Second Party") are not purchasing the property but paying the "First Party" for possession and use of the property, as well as the potential future right to become the owner of the property. YOU WILL NOT OWN THIS PROPERTY UPON SIGNING THIS AGREEMENT OR TAKING POSSESSION OF THE PROPERTY. This transaction is similar to a seller financing transaction; however, you will not take title to the property until you fulfill all terms of the Agreement for Deed including but not limited to making all timely payments as set forth therein. In that regard, if you breach any obligations you have in the Agreement including default of any payments owed therein, '), r(d.seller_name, { bold: true }), r(' (the "First Party") will have the right to terminate this Agreement and treat you as a tenant. That means you could face potential eviction and will lose your right to become an owner of this property. Given the nature and complexity of this transaction, you are strongly encouraged and advised to seek independent legal advice regarding this Agreement.')]),
    blank(),
    p([r('Total monthly payment is '), r(d.total_monthly_payment || '', { bold: true }), r(', this includes property taxes.')]),
    blank(),
    p([r('The Parties have hereunto set their hands and seals the day and year first above written.')]),
    blank(),
    p([r('"FIRST PARTY"', { bold: true })]),
    blank(),
    p([r('________________________________________')]),
    p([r(d.seller_name)]),
    p([r('Signature')]),
    blank(),
    p([r('"SECOND PARTY"', { bold: true })]),
    blank(),
  ];

  // Add up to 4 buyer signature rows
  for (let i = 0; i < buyers.length; i++) {
    const email = d['buyer' + (i+1) + '_email'] || '';
    const phone = d['buyer' + (i+1) + '_phone'] || '';
    children.push(p([r('________________________________________')]));
    children.push(p([r(buyers[i] + (email ? '   Email: ' + email : '') + (phone ? '   Phone: ' + phone : ''))]));
    children.push(p([r('Signature                                      Printed Name')]));
    children.push(blank());
  }

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{
      properties: { page: { size: LEGAL, margin: MARGIN } },
      children
    }]
  });
}

function buildBuyerAcknowledgments(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const servicing = d.servicing_fee || '$120.00';
  const totalPITI = d.total_monthly_payment || '_______________';

  const children = [
    heading('Important Buyer Acknowledgments \u2013 ' + d.property_address),
    blank(),
    p([r('_____ I understand that my monthly payment of '), r(totalPITI, { bold: true }), r(' includes principal, interest, estimated taxes, insurance, and a '), r(servicing, { bold: true }), r(' servicing fee. This amount may change if those costs increase, with proper notice.')]),
    blank(),
    p([r('_____ I understand that payments are due on the 1st of each month and are considered late after the 5th. A 10% late fee will apply if payment is not received by the 5th.')]),
    blank(),
    p([r('_____ I understand that if I default, the Seller will give me written notice and I will have 15 days to fix the issue. If I don\u2019t, the Seller may declare the full remaining balance due. If the default continues for 45 days or more, the Seller may pursue either forfeiture or foreclosure under applicable law, and I may be responsible for court costs and administrative fees.')]),
    blank(),
    p([r('_____ I am purchasing this property '), r('as-is', { bold: true }), r(' and have had the opportunity to conduct my own due diligence or inspections. '), r(d.seller_name, { bold: true }), r(' has made no representations regarding the condition of the property.')]),
    blank(),
    p([r('_____ I am responsible for maintaining the property, including all required repairs, lawn care, utilities, working smoke detectors at all times, and insurance, with '), r(d.seller_name, { bold: true }), r(' listed as an additional insured.')]),
    blank(),
    p([r('_____ I understand that '), r(d.seller_name, { bold: true }), r(' is not escrowing taxes or insurance and is not acting in an escrow or trust capacity.')]),
    blank(),
    p([r('_____ I understand that the seller has done NO inspections and has no knowledge of the condition of the house including, but not limited to plumbing, HVAC, Electrical, Roof, etc, and that any work needed will be '), r('solely', { bold: true }), r(' at the buyers expense.')]),
    blank(),
    p([r('_____ I agree to assume full responsibility for obtaining a valid certificate of occupancy, if required by local regulations, prior to the property being occupied. This process may involve engaging a licensed contractor to pull the necessary permits and conduct essential repairs.')]),
    blank(),
    p([r('_____ I shall transfer utilities into my name within 48 hours of executing this agreement.')]),
    blank(),
    p([r('_____ I have received the amortization schedule.')]),
    blank(),
    p([r('_____ I have received the following required documents:', { bold: true })]),
    p([r('- Seller Disclosure Statement')], { indent: 360 }),
    p([r('- Lead-Based Paint Disclosure')], { indent: 360 }),
    p([r('- Radon Disclosure')], { indent: 360 }),
    p([r('- AS IS Sale Disclosure')], { indent: 360 }),
    p([r('- Deposit Agreement')], { indent: 360 }),
    p([r('- Amortization Schedule')], { indent: 360 }),
    p([r('- Truth In Lending Disclosure')], { indent: 360 }),
    p([r('- Agreement for Deed')], { indent: 360 }),
    blank(),
    hRule(),
  ];

  buyers.forEach((b, i) => {
    const email = d['buyer' + (i+1) + '_email'] || '';
    const phone = d['buyer' + (i+1) + '_phone'] || '';
    children.push(p([r('Buyer Signature: ________________________________________    Date: ________________')]));
    children.push(p([r('Printed Name: '), r(b, { bold: true })]));
    if (email) children.push(p([r('Email: '), r(email, { bold: true }), r('    Phone: '), r(phone, { bold: true })]));
    children.push(blank());
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

function buildTIL(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const border = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  const borders = { top: border, bottom: border, left: border, right: border };
  function cell(text, bold, w) {
    return new TableCell({
      borders, width: { size: w, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [p([r(text, { bold: bold || false, size: 20 })])]
    });
  }

  const children = [
    heading('Federal Truth-in-Lending Disclosure Statement'),
    p([r('This disclosure is provided in compliance with the Federal Truth-in-Lending Act. Please review carefully before signing.')]),
    blank(),
    p([r('LOAN INFORMATION:', { bold: true })]),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [3200, 6880],
      rows: [
        new TableRow({ children: [cell('Borrower(s) Name(s):', true, 3200), cell(buyers.join(', '), false, 6880)] }),
        new TableRow({ children: [cell('Lender/Seller Name:', true, 3200), cell(d.seller_name, false, 6880)] }),
        new TableRow({ children: [cell('Property Address:', true, 3200), cell(d.property_address, false, 6880)] }),
        new TableRow({ children: [cell('Loan Amount:', true, 3200), cell(d.loan_amount_numbers, false, 6880)] }),
        new TableRow({ children: [cell('Loan Term (months/years):', true, 3200), cell('360/30', false, 6880)] }),
        new TableRow({ children: [cell('Date of Transaction:', true, 3200), cell(d.doc_date, false, 6880)] }),
      ]
    }),
    blank(),
    p([r('KEY DISCLOSURES:', { bold: true })]),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [4500, 5580],
      rows: [
        new TableRow({ children: [cell('Annual Percentage Rate:', true, 4500), cell(d.interest_rate_numbers + '%', false, 5580)] }),
        new TableRow({ children: [cell('Finance Charge (Total cost of credit):', true, 4500), cell(d.finance_charge || '', false, 5580)] }),
        new TableRow({ children: [cell('Amount Financed (Credit provided):', true, 4500), cell(d.loan_amount_numbers, false, 5580)] }),
        new TableRow({ children: [cell('Total of Payments (Amount borrower will have paid after all payments are made):', true, 4500), cell(d.total_of_payments || '', false, 5580)] }),
      ]
    }),
    blank(),
    p([r('Payment Schedule', { bold: true })]),
    p([r('Number of Payments: 360   \u2014   Amount of Each Payment: '), r(d.pi_numbers, { bold: true })]),
    blank(),
    p([r('Other Terms and Conditions', { bold: true })]),
    p([r('Late Payment Charge: 10% of total monthly payment')]),
    p([r('Prepayment:  ( X )  Allowed without penalty     (   ) Penalty Applies')]),
    p([r('Security: This loan is secured by the property described above')]),
    blank(),
    p([r('Acknowledgement of Receipt', { bold: true })]),
    p([r('You acknowledge receiving a complete copy of this disclosure before becoming obligated on the loan.')]),
    blank(),
    hRule(),
  ];

  buyers.forEach(b => {
    children.push(p([r('Borrower Signature: ____________________________    Date: ________________')]));
    children.push(p([r('Printed Name: '), r(b, { bold: true })]));
    children.push(blank());
  });
  children.push(p([r('Lender/Seller Signature: ____________________________    Date: ________________')]));

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

function buildBuyerContactInfo(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('Buyer Contact & Emergency Information'),
    blank(),
    p([r('Property Address: '), r(d.property_address, { bold: true })]),
    blank(),
  ];
  buyers.forEach((b, i) => {
    const email = d['buyer' + (i+1) + '_email'] || '';
    const phone = d['buyer' + (i+1) + '_phone'] || '';
    children.push(p([r('Buyer ' + (i+1) + ' Name: '), r(b, { bold: true })]));
    children.push(p([r('Phone: '), r(phone || '_______________________'), r('     Email: '), r(email || '_______________________')]));
    children.push(blank());
  });
  children.push(p([r('Mailing Address (if different from property): _____________________________________________')]));
  children.push(blank());
  children.push(p([r('Emergency Contact Name & Phone: _________________________________________________________')]));

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

function buildDepositAgreement(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const buyerStr = buyers.join(' and ');
  const children = [
    heading('DEPOSIT AGREEMENT - For Agreement for Deed/Installment Sale'),
    blank(),
    p([r('This Non-Refundable Deposit Agreement ("Agreement") is made as of '), r(d.doc_date, { bold: true }), r(' between the Seller '), r(d.seller_name, { bold: true }), r(' and Buyer(s) '), r(buyerStr, { bold: true })]),
    blank(),
    p([r('The property that is the subject of this Agreement is commonly known as:')]),
    p([r('Property Address: '), r(d.property_address, { bold: true })]),
    blank(),
    p([r('Underlying Agreement for Deed. The parties intend to enter into an Beneficial Interest in a Trust/Installment Sale Agreement ("Agreement for Deed") for the property described above.')]),
    blank(),
    p([r('Deposit. Buyer agrees to pay a non-refundable deposit (the "Deposit") to Seller as follows: '), r(d.deposit_amount || '________________', { bold: true }), r('. Unless otherwise stated in this Agreement, the Deposit is non-refundable and will be applied to the purchase price of the Property at Closing.')]),
    blank(),
    p([r('Closing Date. Buyer agrees to complete the purchase of the Property and close under the Agreement for Deed on or before: Closing Deadline: '), r(d.closing_date || '________________', { bold: true }), r(' (Closing Date). Time is of the essence with respect to Buyer\u2019s obligation to close on or before the Closing Date.')]),
    blank(),
    p([r('Failure to Close / Liquidated Damages. If Buyer fails to close the purchase of the Property on or before the Closing Date, and such failure is not caused by Seller\u2019s default of inability to perform:')]),
    p([r('Buyers Deposit shall be fully earned by Seller and non-refundable; Buyer acknowledges and agrees that the Deposit shall serve as liquidated damages, not as a penalty; and Upon Sellers receipt and retention of the Deposit, Seller shall be released from any further obligation to sell the Property to Buyer, unless the parties agree otherwise in writing. Buyer understands and agrees that Seller may market or sell the Property to another buyer after Buyers failure to close by the Closing Date.')], { indent: 360 }),
    blank(),
    p([r('Representations. Buyer represents that Buyer has had the opportunity to inspect the Property, or have it inspected; and Buyer is satisfied with the condition of the Property or is willing to accept it "as-is, where-is", subject to the terms of the Agreement for Deed.')]),
    blank(),
    p([r('Governing Law: This agreement shall be governed by and construed in accordance with the State of '), r(d.state, { bold: true }), r('.')]),
    blank(),
    p([r('Entire Agreement: This Agreement contains the entire understanding between parties regarding the Deposit and the Closing Date described above and supplements (and does not replace) the Beneficial Interest in a Trust. In the event of a conflict between this Agreement and the Beneficial Interest in a Trust regarding the Deposit, the terms of this Agreement shall control, unless the parties agree in writing otherwise.')]),
    blank(),
    p([r('Electronic Signatures: Signatures transmitted by electronic means (including scanned or digitally signed copies) shall be deemed original signatures and fully binding on the parties.')]),
    blank(),
    hRule(),
    p([r('__________________________________________')]),
    p([r(d.seller_name + '   \u2014   Seller')]),
    p([r('Date: ________________')]),
    blank(),
  ];
  buyers.forEach(b => {
    children.push(p([r('__________________________________________')]));
    children.push(p([r(b + '   \u2014   Buyer')]));
    children.push(p([r('Date: ________________')]));
    children.push(blank());
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

function buildInsuranceDisclosure(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const buyerStr = buyers.join(', ');
  const children = [
    heading('PROPERTY INSURANCE DISCLOSURE AND WAIVER OPTION'),
    blank(),
    p([r('This Property Insurance Disclosure and Waiver Option ("Disclosure") is provided to '), r(buyerStr, { bold: true }), r(' ("Buyer(s)") in connection with the Contract for Deed ("Agreement") for the property located at '), r(d.property_address, { bold: true }), r(' ("Property") with '), r(d.seller_name, { bold: true }), r(' ("Seller").')]),
    p([r('The purpose of this Disclosure is to outline the insurance requirements for the Property and provide an option for the Buyer(s) to waive maintaining their own property insurance policy under specific conditions.')]),
    blank(),
    sectionHead('Insurance Requirement'),
    p([r('The Buyer(s) shall maintain, at their sole expense, a property insurance policy ("Policy") covering the Property for the duration of the Agreement. The Policy must provide coverage for all risks of physical loss or damage to the Property, including but not limited to fire, theft, vandalism, and natural disasters, in an amount of the purchase price or sufficient to cover the replacement cost of the property (whichever is higher), as well as liability insurance with at least $100,000 coverage.')]),
    p([r('The policy must name '), r(d.seller_name, { bold: true }), r(', as an additional interest and loss payee. Proof of insurance, including a certificate of insurance or a copy of the Policy, must be provided to Seller within 3 Calendar days of the execution of this Agreement and annually thereafter upon policy renewal. Failure to maintain the required insurance or provide proof of coverage may result in a default under the Agreement, at Seller\u2019s discretion.')]),
    blank(),
    sectionHead('Option to Waive Buyer\'s Property Insurance'),
    p([r('The Buyer(s) may elect to waive the requirement to maintain their own property insurance policy, provided they agree to the following terms:')]),
    p([r('Waiver Fee: In lieu of maintaining their own property insurance policy, Buyer(s) shall pay a monthly fee of $125.00 ("Waiver Fee") to Seller, which will be included in the Buyer(s)\u2019 monthly house payment. This fee is intended to offset the cost of property insurance maintained by Seller on the Property.')], { indent: 360 }),
    p([r('Adjustment of Waiver Fee: The Waiver Fee is subject to adjustment at Seller\u2019s discretion based on changes in the cost of the insurance policy maintained by Seller. Buyer(s) will be notified in writing of any adjustment to the Waiver Fee at least 30 Calendar days prior to the effective date of the change.')], { indent: 360 }),
    p([r('No Guarantee of Exact Cost: The Waiver Fee does not necessarily reflect the exact cost of the insurance policy maintained by Seller and may differ from the actual premium paid by Seller.')], { indent: 360 }),
    p([r('Deductible Responsibility: The property insurance policy maintained by Seller has a deductible of $5,000. If Buyer(s) elect to waive maintaining their own property insurance, they shall be responsible for paying the first $5,000 of any insurance claim filed under the Seller\u2019s policy for damage to the Property.')], { indent: 360 }),
    p([r('Limited Coverage: The insurance policy maintained by Seller covers only the physical structure of the Property and does not include coverage for personal liability or personal property/contents of the Buyer(s). If Buyer(s) elect to waive maintaining their own property insurance, they are required to obtain and maintain a renters insurance policy to cover personal liability and personal property/contents at their sole expense. Proof of renters insurance must be provided to Seller within 5 Calendar days of electing this waiver and annually thereafter upon policy renewal.')], { indent: 360 }),
    blank(),
    sectionHead('Acknowledgment and Election'),
    p([r('By signing below, the Buyer(s) acknowledge that they have read and understand the insurance requirements and waiver option outlined in this Disclosure. The Buyer(s) elects to proceed as follows (check one):')]),
    blank(),
    p([r('[  ] Maintain Own Property Insurance: The Buyer(s) will obtain and maintain a property insurance policy meeting the requirements outlined in Section 1, naming '), r(d.seller_name, { bold: true }), r(', as an additional interest and loss payee.')]),
    blank(),
    p([r('[  ] Waive Property Insurance and Pay Waiver Fee: The Buyer(s) elect to waive maintaining their own property insurance policy and agree to pay the Waiver Fee of $110.00 per month, subject to adjustment, as part of their monthly house payment. Buyer(s) further agrees to be responsible for the first $5,000 of any insurance claim under Seller\u2019s policy and to obtain and maintain a renters insurance policy to cover personal liability and personal property/contents, as outlined in Section 2.')]),
    blank(),
    sectionHead('Default and Remedies'),
    p([r('Failure to comply with the insurance requirements or the terms of the waiver option (including payment of the Waiver Fee, responsibility for the $5,000 deductible, and maintaining renters\u2019 insurance) may constitute a default under the Agreement. In such cases, Seller reserves the right to pursue all available remedies, including but not limited to obtaining insurance on behalf of the Buyer(s) and adding the cost to the Buyer(s)\u2019 obligations under the Agreement, or terminating the Agreement in accordance with its terms.')]),
    blank(),
    sectionHead('Buyer(s) Acknowledgment'),
    p([r('By signing below, the Buyer(s) acknowledge that they have been informed of their obligations under this Disclosure and agree to comply with the selected option. The Buyer(s) further acknowledge that they have been advised to consult with an insurance professional to ensure adequate coverage for their needs.')]),
    blank(),
    hRule(),
  ];
  buyers.forEach((b, i) => {
    const email = d['buyer' + (i+1) + '_email'] || '';
    const phone = d['buyer' + (i+1) + '_phone'] || '';
    children.push(p([r('Buyer Signature: ____________________________    Date: ________________')]));
    children.push(p([r(b + (email ? '   |   ' + email : '') + (phone ? '   |   ' + phone : ''), { bold: true })]));
    children.push(blank());
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN } }, children }]
  });
}

function buildAsIsAddendum(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('AS-IS CONDITION & DUE DILIGENCE ADDENDUM'),
    blank(),
    p([r('This As-Is Condition and Due Diligence Addendum ("Addendum") is incorporated into and made part of the Agreement for Deed / Seller-Financing Agreement between the undersigned Buyer ("Second Party") and Seller ("First Party") regarding the subject property located at '), r(d.property_address, { bold: true }), r('.')]),
    blank(),
    sectionHead('1. AS-IS SALE'),
    p([r('Buyer acknowledges and agrees that the Property is being sold strictly in its current AS-IS, WHERE-IS condition with all faults, whether known or unknown. Seller makes no warranties or representations, express or implied, regarding condition, habitability, value, or future performance.')]),
    blank(),
    sectionHead('2. SELLER NON-OCCUPANCY & LIMITED KNOWLEDGE'),
    p([r('Buyer acknowledges Seller is an investor who has never occupied the Property and may have never personally visited it. Seller has limited or no first-hand knowledge of condition.')]),
    blank(),
    sectionHead('3. NO RELIANCE'),
    p([r('Buyer confirms they are not relying on statements, marketing materials, or opinions from Seller or agents. Buyer relies solely on independent investigations.')]),
    blank(),
    sectionHead('4. BUYER DUE DILIGENCE'),
    p([r('Buyer is strongly encouraged to obtain inspections including general, structural, pest, environmental, and contractor evaluations. All inspections are Buyer\u2019s responsibility.')]),
    blank(),
    sectionHead('5. ACCEPTANCE OF CONDITION'),
    p([r('Buyer affirms adequate opportunity to inspect and fully accepts the Property condition. All repairs and improvements are Buyer\u2019s responsibility.')]),
    blank(),
    sectionHead('6. RELEASE'),
    p([r('Buyer releases Seller from any claims related to condition, defects, code issues, or repairs discovered after closing or possession.')]),
    blank(),
    sectionHead('7. VOLUNTARY AGREEMENT'),
    p([r('Buyer confirms this Addendum is signed freely and is a material part of the purchase agreement.')]),
    blank(),
    hRule(),
  ];
  buyers.forEach((b, i) => {
    const email = d['buyer' + (i+1) + '_email'] || '';
    const phone = d['buyer' + (i+1) + '_phone'] || '';
    children.push(p([r('Buyer Initials: ________    Date: ________________')]));
    children.push(p([r('Buyer Signature: ____________________________    Date: ________________')]));
    children.push(p([r(b + (email ? '   |   ' + email : '') + (phone ? '   |   ' + phone : ''), { bold: true })]));
    children.push(blank());
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

function buildLeadPaint(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('Disclosure of Information on Lead-Based Paint and/or Lead-Based Paint Hazards'),
    blank(),
    p([r('Lead Warning Statement', { bold: true, underline: true })]),
    p([r('Every purchaser of any interest in residential real property on which a residential dwelling was built prior to 1978 is notified that such property may present exposure to lead from lead-based paint that may place young children at risk of developing lead poisoning. Lead poisoning in young children may produce permanent neurological damage, including learning disabilities, reduced intelligence quotient, behavioral problems, and impaired memory. Lead poisoning also poses a particular risk to pregnant women. The seller of any interest in residential real property is required to provide the buyer with any information on lead-based paint hazards from risk assessments or inspections in the seller\'s possession and notify the buyer of any known lead-based paint hazards. A risk assessment or inspection for possible lead-based paint hazards is recommended prior to purchase.'), ], { italics: false }),
    blank(),
    p([r('Property Address: '), r(d.property_address, { bold: true, underline: true })]),
    blank(),
    p([r('Seller\'s Disclosure (Initial)', { bold: true })]),
    p([r('(a) Presence of lead-based paint and/or lead-based paint hazards (check (i) or (ii) below):')]),
    p([r('[  ] Known lead-based paint and/or lead-based paint hazards are present in the housing (explain).')], { indent: 720 }),
    p([r('[X] Seller has no knowledge of lead-based paint and/or lead-based paint hazards in the housing.')], { indent: 720 }),
    p([r('(b) Records and reports available to the seller (check (i) or (ii) below):')]),
    p([r('[  ] (i) Seller has provided the purchaser with all available records and reports pertaining to lead-based paint and/or lead-based paint hazards in the housing (list documents below).')], { indent: 720 }),
    p([r('[X] (ii) Seller has no reports or records pertaining to lead-based paint and/or lead-based paint hazards in the housing.')], { indent: 720 }),
    blank(),
    p([r('Purchaser\'s Acknowledgment (initial)', { bold: true })]),
    p([r('_N/A_ (c) Purchaser has received copies of all information listed above.')]),
    p([r('___/___  (d) Purchaser has received the pamphlet '), r('Protect Your Family from Lead in Your Home', { italics: true }), r('.')]),
    p([r('___/___  (e) Purchaser has (check (i) or (ii) below):')]),
    p([r('[  ] received a 10-day opportunity (or mutually agreed upon period) to conduct a risk assessment or inspection for the presence of lead-based paint and/or lead-based paint hazards; or')], { indent: 720 }),
    p([r('[X] waived the opportunity to conduct a risk assessment or inspection for the presence of lead-based paint and/or lead-based paint hazards.')], { indent: 720 }),
    blank(),
    p([r('Agent\'s Acknowledgment (initial)', { bold: true })]),
    p([r('(f) _N/A_ Agent has informed the seller of the seller\'s obligations under 42 U.S.C. 4852d and is aware of his/her responsibility to ensure compliance.')]),
    blank(),
    p([r('Certification of Accuracy'), r(' The following parties have reviewed the information above and certify, to the best of their knowledge, that the information they have provided is true and accurate.', { size: 19 })], {}),
    blank(),
    hRule(),
    p([r('Seller: ____________________________    Date: ________________')]),
    blank(),
  ];
  buyers.forEach(b => {
    children.push(p([r('Buyer: ____________________________    Date: ________________')]));
    children.push(p([r(b, { bold: true })]));
    children.push(blank());
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

function buildRadonDisclosure(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const children = [
    heading('DISCLOSURE OF INFORMATION ON RADON HAZARDS'),
    p([r('(For Current and Prospective Buyers)')], { align: AlignmentType.CENTER }),
    blank(),
    p([r('Radon Warning Statement', { bold: true, underline: true })]),
    p([r('Each buyer in this residence or dwelling unit is notified that the property may present exposure to levels of indoor radon gas that may place the occupants at risk of developing radon-induced lung cancer. Radon, a Class-A human carcinogen, is the leading cause of death in private homes and the leading cause of lung cancer in nonsmokers. The seller of any residence is required to provide each tenant with any information on radon test results of the dwelling unit that present a radon hazard to the buyer.')]),
    p([r('The Illinois Emergency Management Agency (IEMA) strongly recommends that ALL properties have a radon test performed and radon hazards mitigated if elevated levels are found in a dwelling unit or a routinely occupied area of a multiple family residence. Elevated radon concentrations can easily be reduced by a radon contractor.')]),
    blank(),
    p([r('Dwelling Unit Address: '), r(d.property_address, { bold: true })]),
    blank(),
    p([r('Seller\'s Disclosure (initial each of the following that apply)', { bold: true })]),
    p([r('(a) ___  Seller has no knowledge of elevated radon concentrations (or records or reports pertaining to elevated radon concentrations) in the dwelling unit.')]),
    p([r('(b) _N/A_  Radon concentrations (at or above the IEMA recommended Radon Action Level 4.0 pCi/L) are known to be present within the dwelling unit.')]),
    p([r('(c) ___  Seller has provided the Buyer with copies of all available records and reports, if any, pertaining to radon concentrations within the dwelling unit.')]),
    blank(),
    p([r('Buyer\'s Acknowledgment (initial each of the following that apply)', { bold: true })]),
    p([r('(d) ___/___  Buyer has received copies of all information listed above.')]),
    p([r('(e) ___/___  Buyer has received the pamphlet "Radon Guide for Tenants".')]),
    blank(),
    p([r('Agent\'s Acknowledgment (initial) (if applicable)', { bold: true })]),
    p([r('(g) _N/A_  Agent has informed the seller of the seller\'s obligations under applicable state law.')]),
    blank(),
    p([r('Certification of Accuracy', { bold: true })]),
    p([r('The following parties have reviewed the information above and each party certifies, to the best of his or her knowledge, that the information he or she provided is true and accurate.')]),
    blank(),
    hRule(),
    p([r('Seller: ____________________________    Date: ________________')]),
    blank(),
  ];
  buyers.forEach(b => {
    children.push(p([r('Buyer: ____________________________    Date: ________________')]));
    children.push(p([r(b, { bold: true })]));
    children.push(blank());
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

function buildMemorandum(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const buyerStr = buyers.join(' and ');
  const children = [
    heading('MEMORANDUM OF AGREEMENT FOR DEED'),
    blank(),
    p([r('THIS IS TO CERTIFY:')]),
    blank(),
    p([r('That an Agreement for Deed between '), r(d.seller_name, { bold: true }), r(', designated "SELLER", and '), r(buyerStr, { bold: true }), r(', designated "BUYER", and escrowed at '), r(d.seller_name, { bold: true }), r(' was entered into on '), r(d.doc_date, { bold: true }), r(', wherein the said Seller agreed to sell and the said Buyer agreed to purchase at the purchase price and in accordance with the terms as set forth in said Agreement the following described real estate commonly known as '), r(d.property_address, { bold: true }), r('.')]),
    blank(),
    p([r('Legal Description: ')]),
    p([r(d.legal_description || '________________________________________')]),
    blank(),
    p([r('P.I.N. #: '), r(d.pid || '________________________________________', { bold: true })]),
    blank(),
    p([r('Subject to easements, restrictions and reservations of former instruments of record if any.')]),
    p([r('Subject to general taxes in the year '), r(new Date().getFullYear().toString(), { bold: true }), r(' payable in '), r((new Date().getFullYear() + 1).toString(), { bold: true }), r(' and thereafter.')]),
    blank(),
    p([r('The undersigned certify further that said Agreement contains among other terms an acknowledgment by Buyer of Buyer\'s Special Warranty Deed reconveying the above described real estate to Seller and authorizing reliance by third parties upon said deed, if recorded, as conclusive evidence of the cancellation of said Agreement and reconveyance of all of Buyer\'s then interest in said real estate to Seller.')]),
    blank(),
    p([r('Dated: ___________________________, 2026')]),
    blank(),
    p([r('Seller: ', { bold: true })]),
    p([r(d.seller_name)]),
    p([r('By:')]),
    p([r(d.signor || d.seller_name), r(' \u2014 Member/Manager')]),
    blank(),
    p([r('SUBSCRIBED AND SWORN to before me this ______ day of _______________, 2026.')]),
    blank(),
    hRule(),
    p([r('BUYER(S):', { bold: true })]),
    blank(),
  ];
  buyers.forEach(b => {
    children.push(p([r('________________________________________')]));
    children.push(p([r(b)]));
    children.push(blank());
    children.push(p([r('SUBSCRIBED AND SWORN to before me this ______ day of _______________, 2026.')]));
    children.push(blank());
    children.push(p([r('Notary Public: ____________________________  (SEAL)')]));
    children.push(blank());
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LETTER, margin: MARGIN } }, children }]
  });
}

function buildSellerDisclosure(d) {
  const buyers = [d.buyer1_name, d.buyer2_name, d.buyer3_name, d.buyer4_name].filter(Boolean);
  const border = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  const borders = { top: border, bottom: border, left: border, right: border };
  function chkCell(text, w) {
    return new TableCell({ borders, width: { size: w, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [p([r(text, { size: 18 })])] });
  }
  function disclosureRow(num, text) {
    return new TableRow({ children: [
      chkCell('', 900), chkCell('NO', 900), chkCell('N/A', 900), chkCell(num + '.  ' + text, 7380)
    ]});
  }

  const items = [
    'Seller has occupied the property within the last 12 months. (If "no," please identify capacity or explain the relationship to property.)',
    'I currently have flood hazard insurance on the property.',
    'I am aware of flooding or recurring leakage problems in the crawl space or basement.',
    'I am aware that the property is located in a floodplain.',
    'I am aware of material defects in the basement or foundation (including cracks and bulges).',
    'I am aware of leaks or material defects in the roof, ceilings, or chimney.',
    'I am aware of material defects in the walls, windows, doors, or floors.',
    'I am aware of material defects in the electrical system.',
    'I am aware of material defects in the plumbing system. (includes such things as water heater, sump pump, water treatment system, sprinkler system, and swimming pool).',
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

  const children = [
    heading('RESIDENTIAL REAL PROPERTY DISCLOSURE REPORT'),
    p([r('NOTICE: THE PURPOSE OF THIS REPORT IS TO PROVIDE PROSPECTIVE BUYERS WITH INFORMATION ABOUT MATERIAL DEFECTS IN THE RESIDENTIAL REAL PROPERTY BEFORE THE SIGNING OF A CONTRACT. THIS REPORT DOES NOT LIMIT THE PARTIES\u2019 RIGHT TO CONTRACT FOR THE SALE OF RESIDENTIAL REAL PROPERTY IN "AS IS" CONDITION. UNDER COMMON LAW, SELLERS WHO DISCLOSE MATERIAL DEFECTS MAY BE UNDER A CONTINUING OBLIGATION TO ADVISE THE PROSPECTIVE BUYERS ABOUT THE CONDITION OF THE RESIDENTIAL REAL PROPERTY EVEN AFTER THE REPORT IS DELIVERED TO THE PROSPECTIVE BUYER. COMPLETION OF THIS REPORT BY THE SELLER CREATES LEGAL OBLIGATIONS ON THE SELLER; THEREFORE THE SELLER MAY WISH TO CONSULT AN ATTORNEY PRIOR TO COMPLETION OF THIS REPORT.'), ], { align: AlignmentType.CENTER }),
    blank(),
    p([r('Property Address: '), r(d.property_address, { bold: true, underline: true })]),
    p([r('Seller\'s Name: '), r(d.seller_name, { bold: true, underline: true })]),
    blank(),
    p([r('This Report is a disclosure of certain conditions of the residential real property listed above in compliance with the Residential Real Property Disclosure Act. This information is provided as of '), r(d.doc_date, { bold: true, underline: true }), r('. The disclosures herein shall not be deemed warranties of any kind by the seller or any person representing any party in this transaction.')]),
    blank(),
    p([r('In this form, "aware" means to have actual notice or actual knowledge without any specific investigation or inquiry. In this form, "material defect" means a condition that would have a substantial adverse effect on the value of the residential real property or that would significantly impair the health or safety of future occupants of the residential real property unless the seller reasonably believes that the condition has been corrected.')]),
    blank(),
    p([r('The seller discloses the following information with the knowledge that even though the statements herein are not deemed to be warranties, prospective buyers may choose to rely on this information in deciding whether or not and on what terms to purchase the residential real property.')]),
    blank(),
    p([r('The seller represents that to the best of his or her actual knowledge, the following statements have been accurately noted as "yes" (correct), "no" (incorrect), or "not applicable" to the property being sold. If the seller indicates that the response to any statement, except number 1, is yes or not applicable, the seller shall provide an explanation in the additional information area of this form.')]),
    blank(),
    new Table({
      width: { size: FULL, type: WidthType.DXA }, columnWidths: [900, 900, 900, 7380],
      rows: [
        new TableRow({ children: [chkCell('YES', 900), chkCell('NO', 900), chkCell('N/A', 900), chkCell('ITEM', 7380)] }),
        ...items.map((item, i) => disclosureRow(i + 1, item))
      ]
    }),
    blank(),
    p([r('If any of the above are marked "not applicable" or "yes", please explain here or use additional pages, if necessary:')]),
    p([r('_________________________________________________________________________')]),
    p([r('_________________________________________________________________________')]),
    blank(),
    p([r('Seller certifies that the seller has prepared this report and certifies that the information provided is based on the actual notice or actual knowledge of the seller without any specific investigation or inquiry on the part of the seller. The seller hereby authorizes any person representing any principal in this transaction to provide a copy of this report, and to disclose any information in the report, to any person in connection with any actual or anticipated sale of the property.')]),
    blank(),
    p([r('THE SELLER ACKNOWLEDGES THAT THE SELLER IS REQUIRED TO PROVIDE THIS DISCLOSURE REPORT TO THE PROSPECTIVE BUYER BEFORE THE SIGNING OF THE CONTRACT AND HAS A CONTINUING OBLIGATION, PURSUANT TO APPLICABLE STATE LAW, TO SUPPLEMENT THIS DISCLOSURE PRIOR TO CLOSING.', { bold: true })]),
    blank(),
    p([r('Seller: ____________________________    Date: ________________')]),
    blank(),
    p([r('THE PROSPECTIVE BUYER IS AWARE THAT THE PARTIES MAY CHOOSE TO NEGOTIATE AN AGREEMENT FOR THE SALE OF THE PROPERTY SUBJECT TO ANY OR ALL MATERIAL DEFECTS DISCLOSED IN THIS REPORT ("AS IS"). THIS DISCLOSURE IS NOT A SUBSTITUTE FOR ANY INSPECTIONS OR WARRANTIES THAT THE PROSPECTIVE BUYER OR SELLER MAY WISH TO OBTAIN OR NEGOTIATE.', { bold: true })]),
    blank(),
    p([r('THE FACT THAT THE SELLER IS NOT AWARE OF A PARTICULAR CONDITION OR PROBLEM IS NO GUARANTEE THAT IT DOES NOT EXIST. THE PROSPECTIVE BUYER IS AWARE THAT THE PROSPECTIVE BUYER MAY REQUEST AN INSPECTION OF THE PREMISES PERFORMED BY A QUALIFIED PROFESSIONAL.', { bold: true })]),
    blank(),
    hRule(),
  ];
  buyers.forEach(b => {
    children.push(p([r('Buyer: ____________________________    Date: ________________')]));
    children.push(p([r(b, { bold: true })]));
    children.push(blank());
  });

  return new Document({
    styles: { default: { document: { run: { font: ARIAL, size: 20 } } } },
    sections: [{ properties: { page: { size: LEGAL, margin: MARGIN } }, children }]
  });
}

// ─── main export ──────────────────────────────────────────────────────────────

module.exports = async function generateDocs(data) {
  const docs = [
    { name: '01_Agreement_for_Deed.docx',        doc: buildAgreementForDeed(data) },
    { name: '02_Buyer_Acknowledgments.docx',      doc: buildBuyerAcknowledgments(data) },
    { name: '03_Truth_in_Lending.docx',           doc: buildTIL(data) },
    { name: '04_Buyer_Contact_Info.docx',         doc: buildBuyerContactInfo(data) },
    { name: '05_Deposit_Agreement.docx',          doc: buildDepositAgreement(data) },
    { name: '06_Insurance_Disclosure.docx',       doc: buildInsuranceDisclosure(data) },
    { name: '07_AS_IS_Addendum.docx',             doc: buildAsIsAddendum(data) },
    { name: '08_Lead_Paint_Disclosure.docx',      doc: buildLeadPaint(data) },
    { name: '09_Radon_Disclosure.docx',           doc: buildRadonDisclosure(data) },
    { name: '10_Memorandum_of_Agreement.docx',    doc: buildMemorandum(data) },
    { name: '11_Seller_Disclosure.docx',          doc: buildSellerDisclosure(data) },
  ];

  const buffers = await Promise.all(docs.map(({ doc }) => Packer.toBuffer(doc)));

  return new Promise((resolve, reject) => {
    const archive = archiver('zip', { zlib: { level: 6 } });
    const chunks = [];
    archive.on('data', chunk => chunks.push(chunk));
    archive.on('end', () => resolve(Buffer.concat(chunks)));
    archive.on('error', reject);
    docs.forEach(({ name }, i) => archive.append(buffers[i], { name }));
    archive.finalize();
  });
};
