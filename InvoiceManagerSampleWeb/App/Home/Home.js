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
                    _document.customXmlParts.addAsync(xml, function (result) { });
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

        });
    }
})();