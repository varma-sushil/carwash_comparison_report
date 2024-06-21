import os
import json

wash_packages_all = []



current_file_path = os.path.dirname(os.path.abspath(__file__))
data_path = os.path.join(current_file_path,"data")
with open(f"{data_path}\done\GetRevenuReportFinancialWashPackage.json","r") as f:
            data =json.load(f)
            
# print(type(data))

data = data.get("data")
# print(type(data))

financialWashPackage = data.get("financialWashPackage")

# print(type(financialWashPackage))

# print(financialWashPackage[0].keys())
# print(financialWashPackage[0])

for wash_package in financialWashPackage:
    # print(wash_package)
    wash_package_structure={}
    # Wash_Packages_ServiceName = wash_package.get("serviceName")
    # Wash_Packages_Unlimited   = wash_package.get("cUnlimited")
    # Wash_Packages_Virtual_Wash = wash_package.get("virtualWashNumber")
    # Wash_Packages_Non_Unlimited  = wash_package.get("nonUnlimited")
    # Wash_Packages_Total          = wash_package.get("total")
    # Wash_Packages_Amount         = wash_package.get("price")
    # Wash_Packages_Total_Amount  = wash_package.get("amount")
    
    wash_package_structure['Wash_Packages_ServiceName'] = wash_package.get("serviceName")
    wash_package_structure["Wash_Packages_Unlimited"]   = wash_package.get("cUnlimited")
    wash_package_structure["Wash_Packages_Virtual_Wash"] = wash_package.get("virtualWashNumber")
    wash_package_structure["Wash_Packages_Non_Unlimited"]  = wash_package.get("nonUnlimited")
    wash_package_structure["Wash_Packages_Total "]   =   wash_package.get("total")
    wash_package_structure["Wash_Packages_Amount"]   =    wash_package.get("price")
    wash_package_structure["Wash_Packages_Total_Amount"]  = wash_package.get("amount")
    # print(type(wash_package.get("amount")))
    wash_packages_all.append(wash_package_structure)
    
    
    
    
# print(json.dumps(wash_packages,indent=4))
#discounts logic
discount_all =[]

with open(f"{data_path}\done\GetRevenuReportFinancialWashDiscounts.json","r") as f:
            data =json.load(f)
            
data = data.get("data")

financialWashDiscounts = data.get("financialWashDiscounts")

for wash_discount in financialWashDiscounts:
    wash_discount_structure = {}
    wash_discount_structure["Wash_Packages_Discount_ServiceName"] = wash_discount.get("discountName")
    wash_discount_structure["Wash_Packages_Discount_Number"]      = wash_discount.get("number")
    wash_discount_structure["Wash_Packages_Discount_Service Price ($)"] = wash_discount.get("discountPrice")
    wash_discount_structure["Wash_Packages_Discount_Total Discount ($)"] = wash_discount.get("totalAmt")
    
    discount_all.append(wash_discount_structure)
    
# print(json.dumps(discount_all,indent=4))


wash_extras_all = []

with open(f"{data_path}\done\GetRevenuReportFinancialPackagesDiscount.json","r") as f:
            data =json.load(f)

data = data.get("data")

financialPackagesDiscount = data.get("financialPackagesDiscount") 

for wash_extra in financialPackagesDiscount:
    wash_extra_structure={}
    
    wash_extra_structure["Wash_Extras_ServiceName"] = wash_extra.get("serviceName")
    wash_extra_structure["Wash_Extras_Number"]      = wash_extra.get("number")
    wash_extra_structure["Wash_Extras_Amount ($)"]  = wash_extra.get("servicePrice")
    wash_extra_structure["Wash_Extras_Total Amount ($)"] = wash_extra.get("totalAmount")
    
    wash_extras_all.append(wash_extra_structure)
    
    
# print(json.dumps(wash_extras_all,indent=4))

unlimited_sales_all =[]

with open(f"{data_path}\done\GetRevenuReportFinancialUnlimitedSales.json","r") as f:
            data =json.load(f)

data = data.get("data")

# print(data.keys())

financialUnlimitedSales  = data.get("financialUnlimitedSales")

for sales_data in financialUnlimitedSales:
    sales_data_structure ={}
    # print(sales_data)
    
    sales_data_structure["Unlimited_Sales"] = sales_data.get("unlimited_Sales")
    sales_data_structure["Unlimited_Sales_Service"] = sales_data.get("serviceName")
    sales_data_structure["Unlimited_Sales_Number"]  = sales_data.get("number")
    sales_data_structure["Unlimited_Sales_Revenue ($)"] = sales_data.get("price")
    
    unlimited_sales_all.append(sales_data_structure)
    

# gift card 

gift_card_sale_all =[]
    

with open(f"{data_path}\done\GetRevenuReportFinancialGiftcardsale.json","r") as f:
            data =json.load(f)

data = data.get("data")

financialGiftcardsale = data.get("financialGiftcardsale")

for gift_card in financialGiftcardsale:
    gift_card_sale_structure = {}
    
    gift_card_sale_structure["GIFT_CARD_REDEEMED_DATE"] = gift_card.get("date")  #giftcarsd sales
    gift_card_sale_structure["GIFT_CARD_REDEEMED_TIME"] = gift_card.get("time")
    gift_card_sale_structure["GIFT_CARD_SALES_Card_Number"] = gift_card.get("coupanNumber")
    gift_card_sale_structure["GIFT_CARD_SALESr_Amount ($)"] = gift_card.get("price")
    gift_card_sale_structure["GIFT_CARD_SALES_Source"]      = gift_card.get("transactionFrom")
    
    gift_card_sale_all.append(gift_card_sale_structure)
    
# print(json.dumps(gift_card_sale_all))


gift_card_sale_all =[]
    

with open(f"{data_path}\done\GetRevenuReportFinancialGiftcardsale.json","r") as f:
            data =json.load(f)

data = data.get("data")

financialGiftcardsale = data.get("financialGiftcardsale")

for gift_card in financialGiftcardsale:
    gift_card_sale_structure = {}
    
    gift_card_sale_structure["GIFT_CARD_REDEEMED_DATE"] = gift_card.get("date")  #giftcarsd sales
    gift_card_sale_structure["GIFT_CARD_REDEEMED_TIME"] = gift_card.get("time")
    gift_card_sale_structure["GIFT_CARD_SALES_Card_Number"] = gift_card.get("coupanNumber")
    gift_card_sale_structure["GIFT_CARD_SALESr_Amount ($)"] = gift_card.get("price")
    gift_card_sale_structure["GIFT_CARD_SALES_Source"]      = gift_card.get("transactionFrom")
    
    gift_card_sale_all.append(gift_card_sale_structure)
    
    
# Discount Discount
discount_discount_all =[]

with open(f"{data_path}\done\GetRevenuReportFinancialWashDiscounts.json","r") as f:
            data =json.load(f)
            
data = data.get("data")

financialWashDiscounts = data.get("financialWashDiscounts")

for wash_discount in financialWashDiscounts:
    discount_discount_structure = {}                                                             # Discount Discount
    discount_discount_structure["DISCOUNTS_Discount"] = wash_discount.get("discountName")
    discount_discount_structure["DISCOUNTS_Number"]      = wash_discount.get("number")
    discount_discount_structure["DISCOUNTS_Price ($)"] = wash_discount.get("discountPrice")
    discount_discount_structure["DISCOUNTS_Revenue"] = wash_discount.get("totalAmt")
    
    discount_discount_all.append(discount_discount_structure)
    
# print(json.dumps(discount_discount_all,indent=4))
    
## Gift card reedemed 

reedemed_giftcard_all =[]

with open(f"{data_path}\done\GetRevenuReportFinancialRevenueSummary.json","r") as f:
            data =json.load(f)
            
data = data.get("data")

financialGiftcardRedeemed = data.get("financialGiftcardRedeemed")

for reedemed_giftcard in financialGiftcardRedeemed:
    reedemed_giftcard_structure = {}
    # print(reedemed_giftcard)
    
    reedemed_giftcard_structure["GIFT_CARD_REDEEMED_DATE"] = reedemed_giftcard.get("date")
    reedemed_giftcard_structure["GIFT_CARD_REDEEMED_TIME"] = reedemed_giftcard.get("time")
    reedemed_giftcard_structure["GIFT_CARD_REDEEMED_CARD_NUMBER"] = reedemed_giftcard.get("coupanNumber")
    reedemed_giftcard_structure["GIFT_CARD_REDEEMED_Amount ($)"]  = reedemed_giftcard.get("price")
    
    reedemed_giftcard_all.append(reedemed_giftcard_structure)
    

# print(json.dumps(reedemed_giftcard_all,indent=4))

# payment

payment_data_all =[]

with open(f"{data_path}\GetRevenuReportFinancialPaymentNew.json","r") as f:
            data =json.load(f)
            
data = data.get('data')
# print(data.keys())

financialPaymentNew = data.get("financialPaymentNew")

for payment in financialPaymentNew:
    payment_structure = {}
    print(payment)

    cash = payment.get("cash")
    creditCard = payment.get("creditCard")
    checkpayment  = payment.get("checkpayment")
    invoiceCustomer = payment.get("invoiceCustomer")
    ach = payment.get("ach")
    
    payment_structure["Payment_Location"] = payment.get("locationName")
    payment_structure["Payment_Cash"]     = cash
    payment_structure["Payment_Credit_Card"]  = creditCard
    payment_structure["Payment_Check"]     = checkpayment
    payment_structure["Payment_Invoice"]   = invoiceCustomer
    payment_structure["Payment_ACH"]       = ach
    payment_structure["Payment_Total ($)"] = sum([cash,creditCard,checkpayment,invoiceCustomer,ach])  ##payment
    
    payment_data_all.append(payment_structure)
    

# print(json.dumps(payment_data_all,indent=4))