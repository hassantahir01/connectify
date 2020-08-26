
function addOrderRows(postData) {
    try {
        var sheet = getSheetByName('orders');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'orders' found in spreadsheet.");
        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        sheet.getRange(newRow, 1).setValue(postData.id);
        sheet.getRange(newRow, 2).setValue(postData.name);
        sheet.getRange(newRow, 3).setValue(postData.email);
        sheet.getRange(newRow, 4).setValue(postData.financial_status);
        sheet.getRange(newRow, 5).setValue(postData.fulfillment_status);
        sheet.getRange(newRow, 6).setValue(postData.subtotal_price);
        sheet.getRange(newRow, 7).setValue(postData.total_shipping_price_set.amount);
        sheet.getRange(newRow, 8).setValue(postData.total_tax);
        sheet.getRange(newRow, 9).setValue(postData.total_price);
        sheet.getRange(newRow, 10).setValue(postData.created_at);
        // customer note
        sheet.getRange(newRow, 14).setValue(postData.note);
        sheet.getRange(newRow, 15).setValue(postData.currency);
        sheet.getRange(newRow, 16).setValue(postData.browser_ip);
        sheet.getRange(newRow, 17).setValue(postData.total_discounts);
        sheet.getRange(newRow, 18).setValue(postData.total_price);

        //loop line items
        var lineItems = postData.line_items;
        for (var i in lineItems) {
            if (i > 0) {
                newRow++;
                sheet.getRange(newRow, 1).setValue(postData.id);
                sheet.getRange(newRow, 2).setValue(postData.name);
                sheet.getRange(newRow, 3).setValue(postData.email);
            }
            sheet.getRange(newRow, 11).setValue(lineItems[i].quantity);
            sheet.getRange(newRow, 12).setValue(lineItems[i].name);
            sheet.getRange(newRow, 13).setValue(lineItems[i].price);


        }
        return ContentService.createTextOutput("Order row added successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}
function updateOrderRows(postData) {
    try {
        var sheet = getSheetByName('orders');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'orders' found in spreadsheet.");
        var foundRowInSheet = false;
        var data = sheet.getDataRange().getValues();
        var firstInstance = false;
        var newRow = 0;
        for (var i = 1; i <= data.length; i++) {
            var currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                if (firstInstance === false) {
                    //exclude header row
                    newRow = i + 1;
                    //  ui.alert('detected row no'+newRow);
                    firstInstance = true;
                    // update the first row found for the line item
                    sheet.getRange(newRow, 1).setValue(postData.id);
                    sheet.getRange(newRow, 2).setValue(postData.name);
                    sheet.getRange(newRow, 3).setValue(postData.email);
                    sheet.getRange(newRow, 4).setValue(postData.financial_status);
                    sheet.getRange(newRow, 5).setValue(postData.fulfillment_status);
                    sheet.getRange(newRow, 6).setValue(postData.subtotal_price);
                    sheet.getRange(newRow, 7).setValue(postData.total_shipping_price_set.amount);
                    sheet.getRange(newRow, 8).setValue(postData.total_tax);
                    sheet.getRange(newRow, 9).setValue(postData.total_price);
                    sheet.getRange(newRow, 10).setValue(postData.created_at);
                    // customer note
                    sheet.getRange(newRow, 14).setValue(postData.note);
                    sheet.getRange(newRow, 15).setValue(postData.currency);
                    sheet.getRange(newRow, 16).setValue(postData.browser_ip);
                    sheet.getRange(newRow, 17).setValue(postData.total_discounts);
                    sheet.getRange(newRow, 18).setValue(postData.total_price);
                    //loop line items
                    var lineItems = postData.line_items;
                    for (var j = 0; j < lineItems.length; j++) {
                        if (j > 0) {
                            sheet.insertRowAfter(newRow);
                            newRow++;
                            ///  ui.alert('inserted row no'+newRow);
                            sheet.getRange(newRow, 1).setValue(postData.id);
                            sheet.getRange(newRow, 2).setValue(postData.name);
                            sheet.getRange(newRow, 3).setValue(postData.email);
                        }
                        sheet.getRange(newRow, 11).setValue(lineItems[j].quantity);
                        sheet.getRange(newRow, 12).setValue(lineItems[j].name);
                        sheet.getRange(newRow, 13).setValue(lineItems[j].price);


                        //delete all rows after this index

                    }
                    newRow++;
                }
                else {
                    //delete the existing rows of same order after updating first row and adding new rows under it for more than one line items
                    sheet.deleteRow(newRow);
                }
            } else {
                //row numbers
                newRow++;
            }
        }
        if (!foundRowInSheet) return addOrderRows(postData);
        return ContentService.createTextOutput("order updated successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}


function updateDelCustomer(postData) {

    try {
        var sheet = getSheetByName('customers');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'customers' found in spreadsheet.");
        var foundRowInSheet = false;
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i <= data.length; i++) {
            var currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                //exclude header row
                var newRow = i + 1;
                // if delete webhook called
                if (postData.request_type == "customers/delete") {
                    sheet.deleteRow(newRow);
                    return ContentService.createTextOutput("Customer row deleted from google sheet.");
                }
                // if update,enable,disable webhook triggered
                sheet.getRange(newRow, 1).setValue(postData.id);
                sheet.getRange(newRow, 2).setValue(postData.first_name);
                sheet.getRange(newRow, 3).setValue(postData.last_name);
                sheet.getRange(newRow, 4).setValue(postData.email);

                if (postData.default_address) {
                    var default_address = postData.default_address;
                    sheet.getRange(newRow, 5).setValue(default_address.company);
                    sheet.getRange(newRow, 6).setValue(default_address.address1);
                    sheet.getRange(newRow, 7).setValue(default_address.address2);
                    sheet.getRange(newRow, 8).setValue(default_address.city);
                    sheet.getRange(newRow, 9).setValue(default_address.province);
                    sheet.getRange(newRow, 10).setValue(default_address.province_code);
                    sheet.getRange(newRow, 11).setValue(default_address.country_name);
                    sheet.getRange(newRow, 12).setValue(default_address.country_code);
                    sheet.getRange(newRow, 13).setValue(default_address.zip);
                }
                sheet.getRange(newRow, 14).setValue(postData.phone);
                sheet.getRange(newRow, 15).setValue(postData.accepts_marketing == 1 ? 'TRUE' : 'FALSE');
                sheet.getRange(newRow, 16).setValue(postData.total_spent);
                sheet.getRange(newRow, 17).setValue(postData.orders_count);
                sheet.getRange(newRow, 18).setValue(postData.tags);
                sheet.getRange(newRow, 19).setValue(postData.note);
                sheet.getRange(newRow, 20).setValue(postData.state);
                return ContentService.createTextOutput("Customer update added successfully.");

            }
        }
        if (!foundRowInSheet) return addCustomer(postData);


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function addCustomer(postData) {

    try {
        var sheet = getSheetByName('customers');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'customers' found in spreadsheet.");
        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        sheet.getRange(newRow, 1).setValue(postData.id);
        sheet.getRange(newRow, 2).setValue(postData.first_name);
        sheet.getRange(newRow, 3).setValue(postData.last_name);
        sheet.getRange(newRow, 4).setValue(postData.email);
        if (postData.default_address) {
            var default_address = postData.default_address;
            sheet.getRange(newRow, 5).setValue(default_address.company);
            sheet.getRange(newRow, 6).setValue(default_address.address1);
            sheet.getRange(newRow, 7).setValue(default_address.address2);
            sheet.getRange(newRow, 8).setValue(default_address.city);
            sheet.getRange(newRow, 9).setValue(default_address.province);
            sheet.getRange(newRow, 10).setValue(default_address.province_code);
            sheet.getRange(newRow, 11).setValue(default_address.country_name);
            sheet.getRange(newRow, 12).setValue(default_address.country_code);
            sheet.getRange(newRow, 13).setValue(default_address.zip);
        }
        sheet.getRange(newRow, 14).setValue(postData.phone);
        sheet.getRange(newRow, 15).setValue(postData.accepts_marketing == 1 ? 'TRUE' : 'FALSE');
        sheet.getRange(newRow, 16).setValue(postData.total_spent);
        sheet.getRange(newRow, 17).setValue(postData.orders_count);
        sheet.getRange(newRow, 18).setValue(postData.tags);
        sheet.getRange(newRow, 19).setValue(postData.note);
        sheet.getRange(newRow, 20).setValue(postData.state);
        return ContentService.createTextOutput("Customer row added successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }


}



function addProduct(postData) {
    try {
        var sheet = getSheetByName('products');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'products' found in spreadsheet.");
        const INDEXED_SHEET_FIELDS = {
            "1": "id", "2": "title", "3": "body_html", "4": "vendor", "5": "product_type",
            "6": "tags", "7": "published_at", "8": "option_1", "10": "option_2", "12": "option_3"
        };
        const VARIANTS_SHEET_FIELDS = {
            "9": "option1", "11": "option2", "13": "option3", "14": "sku", "15": "gram", "16": "inventory_quantity",
            "17": "fulfillment_service", "18": "price", "19": "compare_at_price", "20": "taxable",
            "21": "barcode", "22": "variant_image", "23": "weight_unit"
        };

        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        for (var keyProduct in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[keyProduct];

            sheet.getRange(newRow, keyProduct).setValue(postData[value]);
        }
        // loop options

        var options = postData.options;

        if (options && options.length) {
            let keyIndex = 8;
            for (var opt_key in options) {
                sheet.getRange(newRow, keyIndex).setValue(options[opt_key].name);
                keyIndex = keyIndex + 2;
            }

        }
        //loop variants
        var variants = postData.variants;

        if (variants && variants.length) {
            for (var opt in variants) {
                if (opt > 0)
                    sheet.getRange(newRow, 1).setValue(postData.id);
                for (var key in VARIANTS_SHEET_FIELDS) {
                    let variant_key_value = VARIANTS_SHEET_FIELDS[key];
                    if (variant_key_value == 'variant_image') {
                        var p_images = postData.images;
                        for (var img_key in p_images) {

                            let current_variant_id = variants[opt].id;
                            let variant_ids = p_images[img_key].variant_ids || [];
                            // ui.alert(current_variant_id);
                            //Logger.log(variant_ids[1]+'==');return;
                            if (variant_ids.indexOf(current_variant_id) != -1) {
                                sheet.getRange(newRow, key).setValue(p_images[img_key].src);
                            }

                        }
                    } else {
                        sheet.getRange(newRow, key).setValue(variants[opt][variant_key_value]);
                    }
                }
                newRow++;
            }
        }

        return ContentService.createTextOutput("Product row added successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function updateProduct(postData) {
    try {
        var sheet = getSheetByName('products');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'products' found in spreadsheet.");
        const INDEXED_SHEET_FIELDS = {
            "1": "id", "2": "title", "3": "body_html", "4": "vendor",
            "5": "product_type", "6": "tags", "7": "published_at", "8": "option_1",
            "10": "option_2", "12": "option_3"
        };
        const VARIANTS_SHEET_FIELDS = {
            "9": "option1", "11": "option2", "13": "option3", "14": "sku", "15": "gram",
            "16": "inventory_quantity", "17": "fulfillment_service", "18": "price",
            "19": "compare_at_price", "20": "taxable", "21": "barcode", "22": "variant_image",
            "23": "weight_unit"
        };
        var foundRowInSheet = false;
        var data = sheet.getDataRange().getValues();
        var firstInstance = false;
        var newRow = 0;
        for (var i = 1; i <= data.length; i++) {
            var currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                if (firstInstance === false) {
                    //exclude header row
                    newRow = i + 1;
                    firstInstance = true;
                    // update the first row found for the line item
                    for (var keyProduct in INDEXED_SHEET_FIELDS) {
                        let value = INDEXED_SHEET_FIELDS[keyProduct];
                        sheet.getRange(newRow, keyProduct).setValue(postData[value]);
                    }
                    // loop options
                    var options = postData.options;
                    if (options && options.length) {
                        let keyIndex = 8;
                        for (var opt_key in options) {
                            sheet.getRange(newRow, keyIndex).setValue(options[opt_key].name);
                            keyIndex = keyIndex + 2;
                        }
                    }
                    //loop variants
                    var variants = postData.variants;
                    if (variants && variants.length) {
                        for (var opt in variants) {

                            if (opt > 0) {
                                sheet.insertRowAfter(newRow);
                                newRow++;
                                sheet.getRange(newRow, 1).setValue(postData.id);
                            }

                            for (var key in VARIANTS_SHEET_FIELDS) {
                                let variant_key_value = VARIANTS_SHEET_FIELDS[key];
                                if (variant_key_value == 'variant_image') {
                                    var p_images = postData.images;
                                    for (var img_key in p_images) {

                                        var current_variant_id = variants[opt].id;
                                        var variant_ids = p_images[img_key].variant_ids || [];
                                        if (variant_ids.indexOf(current_variant_id) != -1) {

                                            sheet.getRange(newRow, key).setValue(p_images[img_key].src);
                                        }

                                    }
                                } else {
                                    sheet.getRange(newRow, key).setValue(variants[opt][variant_key_value]);
                                }
                            }

                        }
                    }
                    newRow++;
                }
                else {

                    //delete the existing rows of same product after updating first row and adding new rows under it for more than one variants

                    sheet.deleteRow(newRow);
                    // return ContentService.createTextOutput(newRow);
                }


            } else {
                //row numbers
                newRow++;
            }


        }

        if (!foundRowInSheet) return addProduct(postData);
        return ContentService.createTextOutput("product updated successfully.");


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}


function updateDraftOrderRows(postData) {
    try {
        var sheet = getSheetByName('draft_orders');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'draft_orders' found in spreadsheet.");
        var foundRowInSheet = false;
        var data = sheet.getDataRange().getValues();
        var firstInstance = false;
        var newRow = 0;
        for (var i = 1; i <= data.length; i++) {
            var currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                if (firstInstance === false) {
                    //exclude header row
                    newRow = i + 1;
                    firstInstance = true;
                    // update the first row found for the line item
                    sheet.getRange(newRow, 1).setValue(postData.id);
                    sheet.getRange(newRow, 2).setValue(postData.name);
                    sheet.getRange(newRow, 3).setValue(postData.email);
                    sheet.getRange(newRow, 4).setValue(postData.status);
                    sheet.getRange(newRow, 5).setValue(postData.currency);
                    if (postData.applied_discount)
                        sheet.getRange(newRow, 6).setValue(postData.applied_discount.amount || 0);
                    sheet.getRange(newRow, 7).setValue(postData.subtotal_price);
                    if (postData.shipping_line)
                        sheet.getRange(newRow, 8).setValue(postData.shipping_line.price || 0);

                    sheet.getRange(newRow, 9).setValue(postData.total_tax);

                    sheet.getRange(newRow, 10).setValue(postData.tax_exempt);

                    sheet.getRange(newRow, 11).setValue(postData.total_price);
                    // note on end
                    sheet.getRange(newRow, 15).setValue(postData.note);

                    //loop line items
                    var lineItems = postData.line_items;
                    for (var j = 0; j < lineItems.length; j++) {

                        if (j > 0) {
                            sheet.insertRowAfter(newRow);
                            newRow++;
                            ///  ui.alert('inserted row no'+newRow);
                            sheet.getRange(newRow, 1).setValue(postData.id);
                            sheet.getRange(newRow, 2).setValue(postData.name);
                        }
                        sheet.getRange(newRow, 12).setValue(lineItems[j].quantity);
                        sheet.getRange(newRow, 13).setValue(lineItems[j].name);
                        sheet.getRange(newRow, 14).setValue(lineItems[j].price);
                        //delete all rows after this index
                    }
                    newRow++;
                }
                else {
                    //delete the existing rows of same order after updating first row and adding new rows under it for more than one line items
                    sheet.deleteRow(newRow);
                }
            } else {
                //row numbers
                newRow++;
            }
        }
        if (!foundRowInSheet)
        //return ContentService.createTextOutput("No record found in draft_orders sheet.");
            return addDraftOrderRows(postData);
        return ContentService.createTextOutput("draft order updated successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function addDraftOrderRows(postData) {
    try {
        var sheet = getSheetByName('draft_orders');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'draft_orders' found in spreadsheet.");
        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        sheet.getRange(newRow, 1).setValue(postData.id);
        sheet.getRange(newRow, 2).setValue(postData.name);
        sheet.getRange(newRow, 3).setValue(postData.email);
        sheet.getRange(newRow, 4).setValue(postData.status);
        sheet.getRange(newRow, 5).setValue(postData.currency);
        if (postData.applied_discount)
            sheet.getRange(newRow, 6).setValue(postData.applied_discount.amount || 0);
        sheet.getRange(newRow, 7).setValue(postData.subtotal_price);
        if (postData.shipping_line)
            sheet.getRange(newRow, 8).setValue(postData.shipping_line.price || 0);

        sheet.getRange(newRow, 9).setValue(postData.total_tax);

        sheet.getRange(newRow, 10).setValue(postData.tax_exempt);

        sheet.getRange(newRow, 11).setValue(postData.total_price);
        // note on end
        sheet.getRange(newRow, 15).setValue(postData.note);

        //loop line items
        var lineItems = postData.line_items;
        for (var j in lineItems) {
            if (j > 0) {
                newRow++;
                sheet.getRange(newRow, 1).setValue(postData.id);
                sheet.getRange(newRow, 2).setValue(postData.name);
            }
            sheet.getRange(newRow, 12).setValue(lineItems[j].quantity);
            sheet.getRange(newRow, 13).setValue(lineItems[j].name);
            sheet.getRange(newRow, 14).setValue(lineItems[j].price);
        }

        SpreadsheetApp.flush();
        return ContentService.createTextOutput("Draft Order row added successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }
}


function addTenderTransactions(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id", "2": "order_id", "3": "amount", "4": "currency", "5": "user_id",
            "6": "test", "7": "processed_at", "8": "remote_reference", "9": "payment_details",
            "11": "payment_method"
        };

        var sheet = getSheetByName('tender_transactions');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'tender_transactions' found in spreadsheet.");

        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];

            if (value == 'payment_details' && postData.payment_details ) {
                sheet.getRange(newRow, 9).setValue(postData.payment_details.credit_card_number || 0);
                sheet.getRange(newRow, 10).setValue(postData.payment_details.credit_card_company || '');
            } else {
                sheet.getRange(newRow, key).setValue(postData[value]);
            }


        }
        return ContentService.createTextOutput("Tender Transactions row added successfully.");

    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }


}


function addCollection(postData) {
    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id", "2": "handle", "3": "title", "4": "updated_at",
            "5": "body_html", "6": "published_at", "7": "sort_order", "8": "template_suffix",
            "9": "published_scope"
        };
        var sheet = getSheetByName('collections');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'collections' found in spreadsheet.");
        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];
            sheet.getRange(newRow, key).setValue(postData[value]);
        }
        return ContentService.createTextOutput("Collection row added successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }


}


function doPost(e) {

    try {
        var postData = JSON.parse(e.postData.contents);
        switch (postData.request_type) {
            case "checkouts/create":
                return addCheckouts(postData);
                break;
            case "checkouts/update":
                return updateCheckouts(postData);
                break;
            case "checkouts/delete":
                return delObjectRows(postData.id, 'checkouts');
                break;

            case "orders/create":
                return addOrderRows(postData);
                break;
            case "orders/delete":
                return delObjectRows(postData.id, 'orders');
                break;
            case "orders/fulfilled":
            case "orders/partially_fulfilled":
            case "orders/paid":
            case "orders/updated":
            case "orders/cancelled":
                //return ContentService.createTextOutput("here order obj successfully.");
                return updateOrderRows(postData);
                break;
            case "customers/create":
                return addCustomer(postData);
                break;
            case "customers/update":
            case "customers/enable":
            case "customers/disable":
            case "customers/delete":
                return updateDelCustomer(postData);
                break;
            case "products/create":
                return addProduct(postData);
                break;
            case "products/update":
                return updateProduct(postData);
                break;
            case "products/delete":
                return delObjectRows(postData.id, 'products');
                break;

            case "draft_orders/create":
                return addDraftOrderRows(postData);
                break;
            case "draft_orders/update":
                return updateDraftOrderRows(postData);
                break;
            case "draft_orders/delete":
                return delObjectRows(postData.id, 'draft_orders');
                break;


            case "collections/create":
                return addCollection(postData);
                break;
            case "collections/update":
            case "collections/delete":
                return updateDelCollection(postData);
                break;
            case "tender_transactions/create":
                return addTenderTransactions(postData);
                break;
            case "carts/create":
                return addCart(postData);
                break;
            case "carts/update":
                return updateCart(postData);
                break;

            case "inventory_items/create":
                return addInventoryItem(postData);
                break;
            case "inventory_items/update":
            case "inventory_items/delete":
                return updDelInventoryItem(postData);
                break;

            case "inventory_levels/connect":
            case "inventory_levels/disconnect":
            case "inventory_levels/update":
                return updateInventoryLevel(postData);
                break;
            case "locations/create":
                return addLocation(postData);
                break;
            case "locations/update":
            case "locations/delete":
                return updDelLocation(postData);
                break;
            case "order_transactions/create":
                return addOrderTransaction(postData);
                break;

            case "fulfillments/create":
                return addFulfillments(postData);
                break;
            case "fulfillments/update":
                return updateFulfillments(postData);
                break;

            case "fulfillment_events/create":
                return addDelFulfillmentEvent(postData);
                break;
            case "shop/update":
                return updateShop(postData);
                break;
            case "themes/create":
            case "themes/update":
            case "themes/publish":
            case "themes/delete":
                return updateTheme(postData);
                break;
            case "customer_groups/create":
            case "customer_groups/update":
            case "customer_groups/delete":
                return cudCustomerGroup(postData);

            case "refunds/create":
                return addRefunds(postData);
                break;
            case "orders/edited":
                return addOrderEditedRows(postData);
                break;
            case "locales/create":
            case "locales/update":
                return addUpdateLocales(postData);
                break;

            default:
                return ContentService.createTextOutput("No function found for this webhook in sheet apps script .");
                break;

        }
    }
    catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}


function updateDelCollection(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id", "2": "handle", "3": "title", "4": "updated_at", "5": "body_html",
            "6": "published_at", "7": "sort_order", "8": "template_suffix", "9": "published_scope"
        };

        var sheet = getSheetByName('collections');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'collections' found in spreadsheet.");
        var data = sheet.getDataRange().getValues();
        var foundRowInSheet = false;

        for (var i = 1; i <= data.length; i++) {
            var  currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                //exclude header row
                var newRow = i + 1;

                // if delete webhook called
                if (postData.request_type == "collections/delete") {
                    sheet.deleteRow(newRow);
                    return ContentService.createTextOutput("Collection row deleted from google sheet.");
                }

                // if update webhook triggered
                for (var key in INDEXED_SHEET_FIELDS) {
                    let value = INDEXED_SHEET_FIELDS[key];
                    sheet.getRange(newRow, key).setValue(postData[value]);
                }
                return ContentService.createTextOutput("Collection updated successfully.");
            }
        }


        if (!foundRowInSheet) return addCollection(postData);


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function addCart(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {"1": "id", "2": "token", "13": "note", "14": "updated_at", "15": "created_at"};
        const LINE_ITEM_SHEET_FIELDS = {
            "3": "quantity",
            "4": "title",
            "5": "discounted_price",
            "6": "line_price",
            "7": "original_line_price",
            "8": "original_price",
            "9": "price",
            "10": "product_id",
            "11": "sku",
            "12": "total_discount"
        };

        var sheet = getSheetByName('cart');

        if (!sheet) return ContentService.createTextOutput("No sheet having name 'cart' found in spreadsheet.");
        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);

        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];
            sheet.getRange(newRow, key).setValue(postData[value]);
        }

        //loop line items
        var lineItems = postData.line_items;

        if (lineItems && lineItems.length) {

            for (var opt in lineItems) {
                if (opt > 0) {
                    sheet.insertRowAfter(newRow);
                    newRow++;
                    sheet.getRange(newRow, 1).setValue(postData.id);
                    sheet.getRange(newRow, 2).setValue(postData.token);
                }

                for (var key in LINE_ITEM_SHEET_FIELDS) {
                    let line_item_key_value = LINE_ITEM_SHEET_FIELDS[key];
                    sheet.getRange(newRow, key).setValue(lineItems[opt][line_item_key_value]);
                }
            }

        }
        return ContentService.createTextOutput("Cart row added successfully.");


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}


function updateCart(postData) {
    try {
        const INDEXED_SHEET_FIELDS = {"1": "id", "2": "token", "13": "note", "14": "updated_at", "15": "created_at"};
        const LINE_ITEM_SHEET_FIELDS = {
            "3": "quantity",
            "4": "title",
            "5": "discounted_price",
            "6": "line_price",
            "7": "original_line_price",
            "8": "original_price",
            "9": "price",
            "10": "product_id",
            "11": "sku",
            "12": "total_discount"
        };

        var sheet = getSheetByName('cart');

        if (!sheet) return ContentService.createTextOutput("No sheet having name 'cart' found in spreadsheet.");
        var foundRowInSheet = false;
        var data = sheet.getDataRange().getValues();
        var firstInstance = false;
        var newRow = 0;

        for (var i = 1; i <= data.length; i++) {
            var  currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                if (firstInstance === false) {
                    //exclude header row
                    newRow = i + 1;
                    firstInstance = true;

                    for (var key in INDEXED_SHEET_FIELDS) {
                        let value = INDEXED_SHEET_FIELDS[key];
                        sheet.getRange(newRow, key).setValue(postData[value]);
                    }

                    //loop line items
                    var lineItems = postData.line_items;

                    if (lineItems && lineItems.length) {

                        for (var opt in lineItems) {
                            if (opt > 0) {
                                sheet.insertRowAfter(newRow);
                                newRow++;
                                sheet.getRange(newRow, 1).setValue(postData.id);
                                sheet.getRange(newRow, 2).setValue(postData.token);
                            }

                            for (var key in LINE_ITEM_SHEET_FIELDS) {
                                let line_item_key_value = LINE_ITEM_SHEET_FIELDS[key];
                                sheet.getRange(newRow, key).setValue(lineItems[opt][line_item_key_value]);
                            }
                        }

                    }
                    newRow++;
                }
                else {
                    //delete the existing rows of same order after updating first row and adding new rows under it for more than one line items
                    sheet.deleteRow(newRow);
                }

            } else {
                //row numbers
                newRow++;
            }


        }

        if (!foundRowInSheet) return addCart(postData);
        return ContentService.createTextOutput("cart updated successfully.");


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function addOrderTransaction(postData) {

    try {
        var sheet = getSheetByName('order_transactions');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'order_transactions' found in spreadsheet.");
// remove uncessary object keys
        delete postData.admin_graphql_api_id;
        delete postData.request_type;
        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        var columnNo = 1;
        sheet.insertRowAfter(lastRow);
        for (var key in postData) {

            let value = postData[key];

            if (key == 'payment_details' && Object.keys(value).length) {

                for (var obj in value) {

                    sheet.getRange(newRow, columnNo).setValue(value[obj] || '');
                    columnNo++;
                }
            }
            else if (key == 'receipt' && Object.keys(value).length) {
                var concatenate_receipts = '';
                for (var obj in value) {
                    concatenate_receipts += (value[obj] || '') + ',';

                }
                sheet.getRange(newRow, columnNo).setValue(concatenate_receipts);
                columnNo++;
            } else {
                sheet.getRange(newRow, columnNo).setValue(value || '');
                columnNo++;
            }


        }
        return ContentService.createTextOutput("Order Transaction added successfully.");

    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }


}

function updateInventoryLevel(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "inventory_item_id",
            "2": "location_id",
            "3": "available",
            "4": "updated_at"
        };
        var sheet = getSheetByName('inventory_level');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'inventory_level' found in spreadsheet.Please create new sheet with this name and keep header fields described into docs.");
        var data = sheet.getDataRange().getValues();
        var foundRowInSheet = false;
        for (var i = 1; i <= data.length; i++) {
            var    currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.inventory_item_id) {
                foundRowInSheet = true;
//exclude header row
                var newRow = i + 1;
// if update webhook triggered
                for (var key in INDEXED_SHEET_FIELDS) {
                    let value = INDEXED_SHEET_FIELDS[key];
                    sheet.getRange(newRow, key).setValue(postData[value]);

                }
                return ContentService.createTextOutput("Item updated successfully.");

            }
        }
        if (!foundRowInSheet) return addInventoryLevel(postData);

    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function addInventoryLevel(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "inventory_item_id",
            "2": "location_id",
            "3": "available",
            "4": "updated_at"
        };
        var sheet = getSheetByName('inventory_level');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'inventory_level' found in spreadsheet.");

        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];
            sheet.getRange(newRow, key).setValue(postData[value]);

        }
        return ContentService.createTextOutput("Item added successfully.");

    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }


}

function updDelInventoryItem(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "sku",
            "3": "created_at",
            "4": "updated_at",
            "5": "requires_shipping",
            "6": "cost",
            "7": "country_code_of_origin",
            "8": "province_code_of_origin",
            "9": "harmonized_system_code",
            "10": "tracked",
            "11": "country_harmonized_system_codes"
        };
        var sheet = getSheetByName('inventory_items');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'inventory_items' found in spreadsheet.Please create new sheet with this name and keep header fields described into docs.");
        var data = sheet.getDataRange().getValues();
        var foundRowInSheet = false;
        for (var i = 1; i <= data.length; i++) {
            var  currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                //exclude header row
                var newRow = i + 1;

                // if delete webhook called
                if (postData.request_type == "inventory_items/delete") {
                    sheet.deleteRow(newRow);
                    return ContentService.createTextOutput("Inventory Item  deleted from google sheet.");
                }

                // if update webhook triggered
                for (var key in INDEXED_SHEET_FIELDS) {
                    let value = INDEXED_SHEET_FIELDS[key];
                    sheet.getRange(newRow, key).setValue(postData[value]);

                }
                return ContentService.createTextOutput("Inventory Item updated successfully.");

            }
        }


        if (!foundRowInSheet)
            return addInventoryItem(postData);


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function addInventoryItem(postData) {
    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "sku",
            "3": "created_at",
            "4": "updated_at",
            "5": "requires_shipping",
            "6": "cost",
            "7": "country_code_of_origin",
            "8": "province_code_of_origin",
            "9": "harmonized_system_code",
            "10": "tracked",
            "11": "country_harmonized_system_codes"
        };
        var sheet = getSheetByName('inventory_items');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'inventory_items' found in spreadsheet.");
        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];
            sheet.getRange(newRow, key).setValue(postData[value]);

        }
        return ContentService.createTextOutput("Inventory Item added successfully.");

    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }


}

function updDelLocation(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "name",
            "3": "address1",
            "4": "address2",
            "5": "city",
            "6": "zip",
            "7": "province",
            "8": "country",
            "9": "phone",
            "10": "created_at",
            "11": "updated_at",
            "12": "country_code",
            "13": "country_name",
            "14": "province_code",
            "15": "legacy",
            "16": "active"
        };
        var sheet = getSheetByName('locations');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'locations' found in spreadsheet.");
        var data = sheet.getDataRange().getValues();
        var foundRowInSheet = false;
        for (var i = 1; i <= data.length; i++) {
            var  currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
//exclude header row
                var newRow = i + 1;

// if delete webhook called
                if (postData.request_type == "locations/delete") {
                    sheet.deleteRow(newRow);
                    return ContentService.createTextOutput("Locations  deleted from google sheet.");
                }

// if update webhook triggered
                for (var key in INDEXED_SHEET_FIELDS) {
                    let value = INDEXED_SHEET_FIELDS[key];
                    sheet.getRange(newRow, key).setValue(postData[value]);
                }
                return ContentService.createTextOutput("Locations updated successfully.");

            }
        }
        if (!foundRowInSheet)
            return  addLocation(postData)
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }
}


function addLocation(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "name",
            "3": "address1",
            "4": "address2",
            "5": "city",
            "6": "zip",
            "7": "province",
            "8": "country",
            "9": "phone",
            "10": "created_at",
            "11": "updated_at",
            "12": "country_code",
            "13": "country_name",
            "14": "province_code",
            "15": "legacy",
            "16": "active"
        };
        var sheet = getSheetByName('locations');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'locations' found in spreadsheet.");

        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];
            sheet.getRange(newRow, key).setValue(postData[value]);

        }
        return ContentService.createTextOutput("Location added successfully.");

    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function getSheetByName(sheet_name) {

    var spread_sheet = SpreadsheetApp.getActive();
    for (var n in spread_sheet.getSheets()) {
        var sheet = spread_sheet.getSheets()[n];
        var name = sheet.getName();
        if (name == sheet_name) {
            return sheet;
        }

    }
    return false;
}

function addDelFulfillmentEvent(postData) {

    try {

        var sheet = getSheetByName('fulfillment_events');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'fulfillment_events' found in spreadsheet.");
        if (postData.request_type == "fulfillment_events/create") {
            delete postData.admin_graphql_api_id;
            delete postData.request_type;
            var lastRow = Math.max(sheet.getLastRow(), 1);
            var newRow = lastRow + 1;
            var columnNo = 1;
            sheet.insertRowAfter(lastRow);
            for (var key in postData) {
                let value = postData[key];
                sheet.getRange(newRow, columnNo).setValue(value || '');
                columnNo++;
            }
            return ContentService.createTextOutput("Full filment event  added successfully.");

        }

        else {
            var data = sheet.getDataRange().getValues();
            var foundRowInSheet = false;
            for (var i = 1; i <= data.length; i++) {
                var currentSheetRow = data[i];
                if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                    foundRowInSheet = true;
//exclude header row
                    var newRow = i + 1;
                    sheet.deleteRow(newRow);
                    return ContentService.createTextOutput("Fulfilment event deleted from google sheet.");
                }
            }
            if (!foundRowInSheet)
                return ContentService.createTextOutput('No row found in sheet with this id');
        }
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }
}


function updateFulfillments(postData) {
    try {
        const INDEXED_SHEET_FIELDS = {

            "1": "id",
            "2": "order_id",
            "3": "status",
            "4": "created_at",
            "5": "service",
            "6": "updated_at",
            "7": "tracking_company",
            "8": "shipment_status",
            "9": "location_id",
            "10": "email",
            "11": "destination",
            "12": "tracking_number",
            "13": "tracking_url",
            "14": "receipt",
            "15": "name"

        };
        const LINE_ITEM_SHEET_FIELDS = {

            "16": "id",
            "17": "title",
            "18": "quantity",
            "19": "sku",
            "20": "line_item_title",
            "21": "vendor",
            "22": "fulfillment_service",
            "23": "requires_shipping",
            "24": "taxable",
            "25": "gift_card",
            "26": "line_item_inventory_management",
            "27": "grams",
            "28": "price",
            "29": "total_discount",
            "30": "fulfillment_status"
        };
        var sheet = getSheetByName('fulfillments');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'fulfillments' found in spreadsheet.");
        var foundRowInSheet = false;
        var data = sheet.getDataRange().getValues();
        var first_instance = false;
        var newRow = 0;
        for (var i = 1; i <= data.length; i++) {
            var   currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                if (first_instance === false) {
//exclude header row
                    newRow = i + 1;
                    first_instance = true;
// update the first row found for the line item

                    for (var key in INDEXED_SHEET_FIELDS) {
                        let value = INDEXED_SHEET_FIELDS[key];
                        sheet.getRange(newRow, key).setValue(postData[value]);
                    }
//loop lineItems
                    var lineItems = postData.line_items;

//loop lineItems
                    var lineItems = postData.line_items;
                    if (lineItems && lineItems.length) {

                        for (var opt in lineItems) {
                            if (opt > 0) {
                                sheet.insertRowAfter(newRow);
                                newRow++;
                                sheet.getRange(newRow, 1).setValue(postData.id);
                                sheet.getRange(newRow, 2).setValue(postData.order_id);
                            }
                            for (var line_item_array_key in LINE_ITEM_SHEET_FIELDS) {
                                let line_item_key_value = LINE_ITEM_SHEET_FIELDS[line_item_array_key];
                                sheet.getRange(newRow, line_item_array_key).setValue(lineItems[opt][line_item_key_value]);
                            }
                        }
                    }
                    newRow++;

                }
                else {
//delete the existing rows of same product after updating first row and adding new rows under it for more than one lineItems
                    sheet.deleteRow(newRow);
// return ContentService.createTextOutput(newRow);
                }
            } else {
//row numbers
                newRow++;
            }
        }


        if (!foundRowInSheet) return addFulfillments(postData);
        return ContentService.createTextOutput("Fullfillment rows updated successfully.");


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function addFulfillments(postData) {
    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "order_id",
            "3": "status",
            "4": "created_at",
            "5": "service",
            "6": "updated_at",
            "7": "tracking_company",
            "8": "shipment_status",
            "9": "location_id",
            "10": "email",
            "11": "destination",
            "12": "tracking_number",
            "13": "tracking_url",
            "14": "receipt",
            "15": "name"

        };
        const LINE_ITEM_SHEET_FIELDS = {
            "16": "id",
            "17": "title",
            "18": "quantity",
            "19": "sku",
            "20": "line_item_title",
            "21": "vendor",
            "22": "fulfillment_service",
            "23": "requires_shipping",
            "24": "taxable",
            "25": "gift_card",
            "26": "line_item_inventory_management",
            "27": "grams",
            "28": "price",
            "29": "total_discount",
            "30": "fulfillment_status"
        };
        var sheet = getSheetByName('fulfillments');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'fulfillments' found in spreadsheet.");


        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
//var ui= SpreadsheetApp.getUi();
        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];

            sheet.getRange(newRow, key).setValue(postData[value]);
        }
//loop lineItems
        var lineItems = postData.line_items;
        if (lineItems && lineItems.length) {

            for (var opt in lineItems) {
                if (opt > 0) {
                    sheet.insertRowAfter(newRow);
                    newRow++;
                    sheet.getRange(newRow, 1).setValue(postData.id);
                    sheet.getRange(newRow, 2).setValue(postData.order_id);
                }
                for (var line_item_array_key in LINE_ITEM_SHEET_FIELDS) {
                    let line_item_key_value = LINE_ITEM_SHEET_FIELDS[line_item_array_key];
                    sheet.getRange(newRow, line_item_array_key).setValue(lineItems[opt][line_item_key_value]);
                }
            }
        }

        return ContentService.createTextOutput("Fullfillment rows added successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}


function updateShop(postData) {

    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "name",
            "3": "email",
            "4": "domain",
            "5": "province",
            "6": "country",
            "7": "address1",
            "8": "zip",
            "9": "city",
            "10": "phone",
            "11": "primary_locale",
            "12": "address2",
            "13": "updated_at",
            "14": "country_code",
            "15": "country_name",
            "16": "currency",
            "17": "customer_email",
            "18": "shop_owner",
            "19": "weight_unit",
            "20": "plan_display_name",
            "21": "plan_name",
            "22": "enabled_presentment_currencies"

        };
        var sheet = getSheetByName('shop');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'shop' found in spreadsheet.Please create new sheet with this name and keep header fields described into docs.");
        var data = sheet.getDataRange().getValues();
        var foundRowInSheet = false;
        for (var i = 1; i <= data.length; i++) {
            var  currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
//exclude header row
                var newRow = i + 1;
// if update webhook triggered
                for (var key in INDEXED_SHEET_FIELDS) {
                    let value = INDEXED_SHEET_FIELDS[key];
                    sheet.getRange(newRow, key).setValue(postData[value]);

                }
                return ContentService.createTextOutput("Shop updated successfully.");

            }
        }
        if (!foundRowInSheet) {

            var lastRow = Math.max(sheet.getLastRow(), 1);
            var newRow = lastRow + 1;
            sheet.insertRowAfter(lastRow);
            for (var key in INDEXED_SHEET_FIELDS) {
                let value = INDEXED_SHEET_FIELDS[key];
                sheet.getRange(newRow, key).setValue(postData[value]);

            }
            return ContentService.createTextOutput("Shop updated successfully.");
        }


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }
}

function updateTheme(postData) {
    try {
        var sheet = getSheetByName('theme');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'theme' found in spreadsheet.Please create new sheet with this name and keep header fields described into docs.");
        var data = sheet.getDataRange().getValues();
        var columnNo = 1;
        delete postData.admin_graphql_api_id;
        var foundRowInSheet = false;
        for (var i = 1; i <= data.length; i++) {
            var currentSheetRow = data[i];
            if (typeof currentSheetRow != 'undefined' && postData.id && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                //exclude header row
                var newRow = i + 1;

                if (postData.request_type == 'themes/delete') {
                    sheet.deleteRow(newRow);
                    return ContentService.createTextOutput("Theme deleted from google sheet.");
                }
                // if update webhook triggered
                delete postData.request_type;
                for (var key in postData) {
                    let value = postData[key];
                    sheet.getRange(newRow, columnNo).setValue(value);
                    columnNo++;

                }
                return ContentService.createTextOutput("Theme updated successfully.");

            }
        }
        if (!foundRowInSheet) {
            var lastRow = Math.max(sheet.getLastRow(), 1);
            var newRow = lastRow + 1;
            delete postData.request_type;
            sheet.insertRowAfter(lastRow);
            for (var keys in postData) {
                let value = postData[keys];
                sheet.getRange(newRow, columnNo).setValue(value);
                columnNo++;

            }
            return ContentService.createTextOutput("Theme added successfully.");
        }
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }
}

function cudCustomerGroup(postData) {
    try {
        var sheet = getSheetByName('customer_saved_search');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'customer_saved_search' found in spreadsheet.Please create new sheet with this name and keep header fields described into docs.");
        var data = sheet.getDataRange().getValues();
        var columnNo = 1;
        var foundRowInSheet = false;
        for (var i = 1; i <= data.length; i++) {
            var currentSheetRow = data[i];
            if (typeof currentSheetRow != 'undefined' && postData.id && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                //exclude header row
                var newRow = i + 1;

                if (postData.request_type == 'customer_groups/delete') {
                    sheet.deleteRow(newRow);
                    return ContentService.createTextOutput("row deleted from google sheet.");
                }
                // if update webhook triggered
                delete postData.request_type;
                for (var key in postData) {
                    let value = postData[key];
                    sheet.getRange(newRow, columnNo).setValue(value);
                    columnNo++;

                }
                return ContentService.createTextOutput("row updated successfully.");

            }
        }
        if (!foundRowInSheet) {
            var lastRow = Math.max(sheet.getLastRow(), 1);
            var newRow = lastRow + 1;
            delete postData.request_type;
            sheet.insertRowAfter(lastRow);
            for (var keys in postData) {
                let value = postData[keys];
                sheet.getRange(newRow, columnNo).setValue(value);
                columnNo++;

            }
            return ContentService.createTextOutput("row added successfully.");
        }
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }
}

function addOrderEditedRows(postData) {
    try {
        var sheet = getSheetByName('order_edits');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'order_edits' found in spreadsheet.Please create new sheet with this name and keep header fields described into docs.");
        var columnNo = 1;

        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        delete postData.request_type;
        sheet.insertRowAfter(lastRow);

        for (var key in postData) {
            let value = postData[key];
            if (key == 'line_items') {
                sheet.getRange(newRow, columnNo).setValue(value.additions);
                columnNo++;
                sheet.getRange(newRow, columnNo).setValue(value.removals);
            } else {
                sheet.getRange(newRow, columnNo).setValue(value);
            }

            columnNo++;

        }
        return ContentService.createTextOutput("Order Edit added successfully.");

    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }
}

function addUpdateLocales(postData) {
    try {
        var sheet = getSheetByName('locales');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'locales' found in spreadsheet.Please create new sheet with this name and keep header fields described into docs.");
        var data = sheet.getDataRange().getValues();
        var columnNo = 1;
        var foundRowInSheet = false;
        delete postData.request_type;
        for (var i = 1; i <= data.length; i++) {
            var currentSheetRow = data[i];
            if (typeof currentSheetRow != 'undefined'  && currentSheetRow[0] == postData.locale) {
                foundRowInSheet = true;
                //exclude header row
                var newRow = i + 1;
                // if update webhook triggered
                for (var key in postData) {
                    let value = postData[key];
                    sheet.getRange(newRow, columnNo).setValue(value);
                    columnNo++;

                }
                return ContentService.createTextOutput("Locale row updated into the sheet successfully.");

            }
        }
        if (!foundRowInSheet) {
            var lastRow = Math.max(sheet.getLastRow(), 1);
            var newRow = lastRow + 1;

            sheet.insertRowAfter(lastRow);
            for (var keys in postData) {
                let value = postData[keys];
                sheet.getRange(newRow, columnNo).setValue(value);
                columnNo++;

            }
            return ContentService.createTextOutput("Locale row added into the sheet successfully.");
        }
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }
}


function addRefunds(postData) {
    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "order_id",
            "3": "created_at",
            "4": "note",
            "5": "processed_at",
            "6": "restock",
            "7": "duties"
        };
        const REFUND_LINE_ITEM_SHEET_FIELDS = {
            "8": "id",
            "9": "quantity",
            "10": "line_item_id",
            "11": "restock_type",
            "12": "subtotal",
            "13": "total_tax",
            "14": "variant_id",
            "15": "title",
            "16": "sku",
            "17": "variant_title",
            "18": "vendor",
            "19": "fulfillment_service",
            "20": "requires_shipping",
            "21": "grams",
            "22": "price",
            "23": "total_discount",
            "24": "fulfillment_status"
        };
        var sheet = getSheetByName('refunds');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'refunds' found in spreadsheet.");
        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        //var ui= SpreadsheetApp.getUi();
        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];

            sheet.getRange(newRow, key).setValue(postData[value]);
        }
        //loop lineItems
        var refundLineItems = postData.refund_line_items;
        if (refundLineItems && refundLineItems.length) {
            for (var opt in refundLineItems) {
                if (opt > 0) {
                    sheet.insertRowAfter(newRow);
                    newRow++;
                    sheet.getRange(newRow, 1).setValue(postData.id);
                    sheet.getRange(newRow, 2).setValue(postData.order_id);
                }
                for (var line_item_array_key in REFUND_LINE_ITEM_SHEET_FIELDS) {
                    let line_item_key_value = REFUND_LINE_ITEM_SHEET_FIELDS[line_item_array_key];
                    if (line_item_array_key >=14)
                    {
                        sheet.getRange(newRow, line_item_array_key).setValue(refundLineItems[opt]['line_item'][line_item_key_value]);
                    }
                    else
                    {
                        sheet.getRange(newRow, line_item_array_key).setValue(refundLineItems[opt][line_item_key_value]);
                    }
                }
            }
        }

        return ContentService.createTextOutput("Refunds row added successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}


function delObjectRows(comparative_col_value, sheet_name) {
    try {

        var sheet = getSheetByName(sheet_name);
        if (!sheet) return ContentService.createTextOutput('No sheet having name ' + sheet_name + ' found in spreadsheet.');
        var foundRowInSheet = false;

        var rows = sheet.getDataRange();
        var numRows = rows.getNumRows();
        var values = rows.getValues();
        var rowsDeleted = 0;

        // delete procedure
        for (var i = 0; i <= numRows - 1; i++) {
            var row = values[i];
            if (typeof row !== 'undefined' && row[0] == comparative_col_value) {
                if (!foundRowInSheet)
                    foundRowInSheet = true;
                sheet.deleteRow((parseInt(i) + 1) - rowsDeleted);
                rowsDeleted++;
            }
        }


        if (!foundRowInSheet) return ContentService.createTextOutput("No record found in sheet.");
        return ContentService.createTextOutput("Object rows deleted successfully.");


    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}


function addCheckouts(postData) {
    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "token",
            "3": "cart_token",
            "4": "email",
            "5": "gateway",
            "6": "created_at",
            "7": "updated_at",
            "8": "note",
            "9": "note_attributes",
            "10": "total_weight",
            "11": "currency",
            "12": "completed_at",
            "13": "phone",
            "14": "source_name",
            "15": "total_discounts",
            "16": "total_line_items_price",
            "17": "total_price",
            "18": "total_tax",
            "19": "subtotal_price",
            "20": "billing_address",
            "21": "shipping_address",
            "22": "customer"


        };
        const LINE_ITEM_SHEET_FIELDS = {
            "23": "variant_id",
            "24": "title",
            "25": "variant_title",
            "26": "variant_price",
            "27": "vendor",
            "28": "sku",
            "29": "grams",
            "30": "gift_card",
            "31": "fulfillment_service",
            "32": "line_price",
            "33": "compare_at_price",
            "34": "price",
            "35": "applied_discounts"

        };
        var sheet = getSheetByName('checkouts');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'checkouts' found in spreadsheet.");


        var lastRow = Math.max(sheet.getLastRow(), 1);
        var newRow = lastRow + 1;
        sheet.insertRowAfter(lastRow);
        //var ui= SpreadsheetApp.getUi();
        for (var key in INDEXED_SHEET_FIELDS) {
            let value = INDEXED_SHEET_FIELDS[key];

            sheet.getRange(newRow, key).setValue(postData[value]);
        }
        //loop lineItems
        var lineItems = postData.line_items;
        if (lineItems && lineItems.length) {

            for (var opt in lineItems) {
                if (opt > 0) {
                    sheet.insertRowAfter(newRow);
                    newRow++;
                    sheet.getRange(newRow, 1).setValue(postData.id);
                }
                for (var line_item_array_key in LINE_ITEM_SHEET_FIELDS) {
                    let line_item_key_value = LINE_ITEM_SHEET_FIELDS[line_item_array_key];
                    sheet.getRange(newRow, line_item_array_key).setValue(lineItems[opt][line_item_key_value]);
                }
            }
        }

        return ContentService.createTextOutput("Checkout row added successfully.");
    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}

function updateCheckouts(postData) {
    try {
        const INDEXED_SHEET_FIELDS = {
            "1": "id",
            "2": "token",
            "3": "cart_token",
            "4": "email",
            "5": "gateway",
            "6": "created_at",
            "7": "updated_at",
            "8": "note",
            "9": "note_attributes",
            "10": "total_weight",
            "11": "currency",
            "12": "completed_at",
            "13": "phone",
            "14": "source_name",
            "15": "total_discounts",
            "16": "total_line_items_price",
            "17": "total_price",
            "18": "total_tax",
            "19": "subtotal_price",
            "20": "billing_address",
            "21": "shipping_address",
            "22": "customer"


        };
        const LINE_ITEM_SHEET_FIELDS = {
            "23": "variant_id",
            "24": "title",
            "25": "variant_title",
            "26": "variant_price",
            "27": "vendor",
            "28": "sku",
            "29": "grams",
            "30": "gift_card",
            "31": "fulfillment_service",
            "32": "line_price",
            "33": "compare_at_price",
            "34": "price",
            "35": "applied_discounts"

        };
        var sheet = getSheetByName('checkouts');
        if (!sheet) return ContentService.createTextOutput("No sheet having name 'checkouts' found in spreadsheet.");
        var foundRowInSheet = false;
        var data = sheet.getDataRange().getValues();
        var first_instance = false;
        var newRow = 0;
        for (var i = 1; i <= data.length; i++) {
            var currentSheetRow = data[i];
            if (typeof currentSheetRow !== 'undefined' && currentSheetRow[0] == postData.id) {
                foundRowInSheet = true;
                if (first_instance === false) {
                    //exclude header row
                    newRow = i + 1;
                    first_instance = true;
                    // update the first row found for the line item

                    for (var key in INDEXED_SHEET_FIELDS) {
                        let value = INDEXED_SHEET_FIELDS[key];
                        sheet.getRange(newRow, key).setValue(postData[value]);
                    }

                    //loop lineItems
                    var lineItems = postData.line_items;
                    if (lineItems && lineItems.length) {

                        for (var opt in lineItems) {
                            if (opt > 0) {
                                sheet.insertRowAfter(newRow);
                                newRow++;
                                sheet.getRange(newRow, 1).setValue(postData.id);
                            }
                            for (var line_item_array_key in LINE_ITEM_SHEET_FIELDS) {
                                let line_item_key_value = LINE_ITEM_SHEET_FIELDS[line_item_array_key];
                                sheet.getRange(newRow, line_item_array_key).setValue(lineItems[opt][line_item_key_value]);
                            }
                        }
                    }
                    newRow++;

                }
                else {
                    //delete the existing rows of same product after updating first row and adding new rows under it for more than one lineItems
                    sheet.deleteRow(newRow);
                    // return ContentService.createTextOutput(newRow);
                }
            } else {
                //row numbers
                newRow++;
            }
        }
        if (!foundRowInSheet) return addCheckouts(postData);
        return ContentService.createTextOutput("checkout rows updated successfully.");

    } catch (err) {
        return ContentService.createTextOutput(err.message);
    }

}







