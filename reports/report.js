$(document).ready(function() {var formatter = new CucumberHTML.DOMFormatter($('.cucumber-report'));formatter.uri("file:src/main/resources/feature/open.feature");
formatter.feature({
  "name": "OpenCartApplication",
  "description": "  In order to create a webpage\n  As a e-commerce vendor\n  I want to get access to the portal",
  "keyword": "Feature"
});
formatter.background({
  "name": "these are the steps common for every scenario to be followed",
  "description": "",
  "keyword": "Background"
});
formatter.step({
  "name": "I have browser with opencartpage",
  "keyword": "Given "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_have_browser_with_opencartpage()"
});
formatter.result({
  "status": "passed"
});
formatter.scenario({
  "name": "ValidRegisterCredentail",
  "description": "",
  "keyword": "Scenario"
});
formatter.step({
  "name": "I enter email as \u0027pravallika12@gmail.com\u0027 and I enter password as \u0027sunny123\u0027",
  "keyword": "When "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_enter_email_as_and_I_enter_password_as(java.lang.String,java.lang.String)"
});
formatter.result({
  "status": "passed"
});
formatter.step({
  "name": "my account should access to the portal with title as \u0027My Account\u0027",
  "keyword": "Then "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.my_account_should_access_to_the_portal_with_title_as(java.lang.String)"
});
formatter.result({
  "status": "passed"
});
formatter.after({
  "status": "passed"
});
formatter.background({
  "name": "these are the steps common for every scenario to be followed",
  "description": "",
  "keyword": "Background"
});
formatter.step({
  "name": "I have browser with opencartpage",
  "keyword": "Given "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_have_browser_with_opencartpage()"
});
formatter.result({
  "status": "passed"
});
formatter.scenario({
  "name": "ClickOnDesktops",
  "description": "",
  "keyword": "Scenario"
});
formatter.step({
  "name": "I click on the product Desktops",
  "keyword": "When "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_click_on_the_product_Desktops()"
});
formatter.result({
  "status": "passed"
});
formatter.step({
  "name": "I should get all Desktops",
  "keyword": "Then "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_should_get_all_Desktops()"
});
formatter.result({
  "status": "passed"
});
formatter.after({
  "status": "passed"
});
formatter.background({
  "name": "these are the steps common for every scenario to be followed",
  "description": "",
  "keyword": "Background"
});
formatter.step({
  "name": "I have browser with opencartpage",
  "keyword": "Given "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_have_browser_with_opencartpage()"
});
formatter.result({
  "status": "passed"
});
formatter.scenario({
  "name": "ClickOnLaptops and Notebooks",
  "description": "",
  "keyword": "Scenario"
});
formatter.step({
  "name": "I click on the product Laptops and Notebooks",
  "keyword": "When "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_click_on_the_product_Laptops_and_Notebooks()"
});
formatter.result({
  "status": "passed"
});
formatter.step({
  "name": "I should get all Laptops and Notebooks",
  "keyword": "Then "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_should_get_all_Laptops_and_Notebooks()"
});
formatter.result({
  "status": "passed"
});
formatter.after({
  "status": "passed"
});
formatter.background({
  "name": "these are the steps common for every scenario to be followed",
  "description": "",
  "keyword": "Background"
});
formatter.step({
  "name": "I have browser with opencartpage",
  "keyword": "Given "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_have_browser_with_opencartpage()"
});
formatter.result({
  "status": "passed"
});
formatter.scenario({
  "name": "ClickOnComponents",
  "description": "",
  "keyword": "Scenario"
});
formatter.step({
  "name": "I click on the product Components",
  "keyword": "When "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_click_on_the_product_Components()"
});
formatter.result({
  "status": "passed"
});
formatter.step({
  "name": "I should get all Components",
  "keyword": "Then "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_should_get_all_Components()"
});
formatter.result({
  "status": "passed"
});
formatter.after({
  "status": "passed"
});
formatter.background({
  "name": "these are the steps common for every scenario to be followed",
  "description": "",
  "keyword": "Background"
});
formatter.step({
  "name": "I have browser with opencartpage",
  "keyword": "Given "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_have_browser_with_opencartpage()"
});
formatter.result({
  "status": "passed"
});
formatter.scenario({
  "name": "ClickOnTablets",
  "description": "",
  "keyword": "Scenario"
});
formatter.step({
  "name": "I click on the product Tablets",
  "keyword": "When "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_click_on_the_product_Tablets()"
});
formatter.result({
  "status": "passed"
});
formatter.step({
  "name": "I should get all Tablets",
  "keyword": "Then "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_should_get_all_Tablets()"
});
formatter.result({
  "status": "passed"
});
formatter.after({
  "status": "passed"
});
formatter.scenarioOutline({
  "name": "ClickOnShopping Cart",
  "description": "",
  "keyword": "Scenario Outline"
});
formatter.step({
  "name": "I enter \u0027\u003cemail\u003e\u0027 and \u0027\u003cpassword\u003e\u0027 click on shopping cart",
  "keyword": "When "
});
formatter.step({
  "name": "I should get the products which I added into the cart",
  "keyword": "Then "
});
formatter.examples({
  "name": "",
  "description": "",
  "keyword": "Examples",
  "rows": [
    {
      "cells": [
        "email",
        "password"
      ]
    },
    {
      "cells": [
        "pravallika12@gmail.com",
        "sunny123"
      ]
    }
  ]
});
formatter.background({
  "name": "these are the steps common for every scenario to be followed",
  "description": "",
  "keyword": "Background"
});
formatter.step({
  "name": "I have browser with opencartpage",
  "keyword": "Given "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_have_browser_with_opencartpage()"
});
formatter.result({
  "status": "passed"
});
formatter.scenario({
  "name": "ClickOnShopping Cart",
  "description": "",
  "keyword": "Scenario Outline"
});
formatter.step({
  "name": "I enter \u0027pravallika12@gmail.com\u0027 and \u0027sunny123\u0027 click on shopping cart",
  "keyword": "When "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_enter_and_click_on_shopping_cart(java.lang.String,java.lang.String)"
});
formatter.result({
  "status": "passed"
});
formatter.step({
  "name": "I should get the products which I added into the cart",
  "keyword": "Then "
});
formatter.match({
  "location": "com.cts.stepdefinitions.StepDefinitions.i_should_get_the_products_which_I_added_into_the_cart()"
});
formatter.result({
  "status": "passed"
});
formatter.after({
  "status": "passed"
});
});