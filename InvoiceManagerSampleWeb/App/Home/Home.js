/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {

        // Checks for the DOM to load.
        $(document).ready(function () {

            var myOrders;
            var _document;

            _document = Office.context.document;
           
            
            setupMyOrders();
            initializeOrder();

            // Make sure the doc that accompanies this sample is loaded.
            // This doc contains the customXMLParts and content controls used by this sample.
            checkSampleDocLoaded();

            function initializeOrder() {
                $.each(myOrders, function (index, value) {
                    $("#orders").append($("<option></option>").val(value.info.id).html(value.info.id));
                });
                $("#orders").change(function () {
                    popOrder($(this).children("option:selected").val());
                });
                $("#populate").click(function () {
                    var selectedOrderID = parseInt($("#orders option:selected").val());
                    _document.customXmlParts.getByNamespaceAsync("", function (result) {
                        if (result.value.length > 0) {
                            for (var i = 0; i < result.value.length; i++) {
                                result.value[i].deleteAsync(function () {
                                });
                            }
                        }
                    });
                    var xml = $.json2xml(findOrder(myOrders, selectedOrderID));
                    _document.customXmlParts.addAsync(xml, function (result) { _document.getSelectedDataAsync(Office.CoercionType.Ooxml);});
                });
                var selOrder = $("#orders option:selected");
                popOrder(selOrder.val());
            }

            function popOrder(o_id) {
                var order = findOrder(myOrders, parseInt(o_id));
                if (order != null) {
                    $("#address_line_1").html(order.customer.address);
                    $("#address_line_2").html(order.customer.address2);
                    $("#date").html(order.info.date);
                    $("#items").html("");
                    $.each(order.items.item, function (index, value) {
                        $("#items").append($("<div></div>").html(value.name + " [" + value.price + "]"));
                    });
                    $("#customer_name").html(order.customer.first_name + " " + order.customer.last_name);
                }
            }

            function findOrder(arr, id) {
                for (var i = 0; i < arr.length; i++) {
                    if (arr[i].info.id === id) {
                        return arr[i];
                    }
                }
                return null;
            }
            function setupMyOrders() {
                myOrders = new Array();
                var order1 = {
                    info: {
                        id: 918291,
                        date: "5/24/2015"
                    },
                    items:
                    {
                        item: [
                        { name: "Diary of a Wussy Kid: Cabin Fever", price: "$20.00" },
                        { name: "1Q95", price: "$16.05" },
                        { name: "A Wild Goose Chase: A Novel", price: "$12.35" },
                        { name: "Kafka in the Woods", price: "$7.86" },
                        { name: "My Isadora", price: "$16.05" },
                        { name: "Sputnik Darling", price: "$11.20" }
                        ]
                    },
                    id: 918291,
                    customer: {
                        first_name: "Lisa",
                        last_name: "Andrews",
                        address: "678 Elm St.",
                        address2: "Redwood City, CA 12202"
                    }
                };
                var order2 = {
                    info: {
                        id: 955847,
                        date: "7/15/2015"
                    },
                    items:
                    {
                        item: [
                        { name: "Transformations 3D", price: "$23.99" },
                        { name: "Deadly Attraction", price: "$9.05" },
                        { name: "Band of Sisters", price: "$59.99" },
                        { name: "Carolina Shore: The Complete First Season", price: "$0.99" },
                        { name: "The Prizefighter", price: "$11.59" },
                        { name: "Swedish Wood", price: "$11.99" }
                        ]
                    },
                    customer: {
                        first_name: "Josh",
                        last_name: "Bailey",
                        address: "345 Main St.",
                        address2: "Woodinville, WA 99901"
                    }
                };
                var order3 = {
                    info: {
                        id: 985412,
                        date: "8/19/2015"
                    },
                    items:
                    {

                        item: [
                        { name: "Shotgun 2: Total War", price: "$53.99" },
                        { name: "Falloff 3", price: "$59.05" },
                        { name: "Call of Need 4", price: "$59.99" },
                        { name: "Tropical 3", price: "$45.99" },
                        { name: "Messy World", price: "$19.59" },
                        { name: "A Wind in the Pillows", price: "$11.99" }
                        ]
                    },
                    customer: {
                        first_name: "Kevin",
                        last_name: "Cook",
                        address: "123 3rd Avenue South",
                        address2: "Seattle, WA 12345"
                    }
                };
                myOrders.push(order1);
                myOrders.push(order2);
                myOrders.push(order3);
            }

            function checkSampleDocLoaded() {

                //Get the URL of the current file.
                Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                    var fileUrl = asyncResult.value.url;

                    // Find the warning section on the page. We'll hide or show it
                    // depending on whether the sample doc is loaded or not. 
                    var elem = document.getElementById('warning');

                    // Note: If you change the name of the sample doc in the InvoiceManager.csproj, don't
                    // forget to update it here 
                    if (fileUrl == "" || fileUrl.indexOf("PackingSlip.docx") == -1) {
                        elem.style.display = 'block'; // show warning
                    }
                    else {
                        elem.style.display = 'none'; // hide warning
                    }
                });
            }

        });
    }
})();

// *********************************************************
//
// Word-Add-in-JavaScript-InvoiceManager, https://github.com/OfficeDev/Word-Add-in-JavaScript-InvoiceManager
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************