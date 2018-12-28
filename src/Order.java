import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;

public class Order {

    private Integer orderId;
    private Integer countProduct;
    private Integer countDay;
    private Integer customerId;
    private Integer managerId;
    private Integer paymentTypeId;
    private Integer deliveryTypeId;
    private Integer statusId;
    private Integer payInStatusId;
//    private Double priceDelivery;
    private Double priceFull;
    private Double payOut;
//    private Double payOther;
    private Double sumAllPays;
    private String customerName;
    private String mail;
    private String phone;
    private String paymentType;
    private String deliveryType;
    private String deliveryAddress;
    private String status;
    private String manager;
    private String description;
    private Date dateIncome;
    private Date dateShipOut;
    private Date dateStart;
    private Date dateShipBack;
    private Date datePayment;
    private Integer discount;
    private Integer productShipStatusId;
    private String productShipStatusDescription;
    private Date dateClose;
    private Date salaryPayDate;

    // расчетывается налету
    private Integer profit;

    public Order() {
    }

    public Order(Integer orderId, Integer countProduct, Integer countDay,
                 Integer customerId, Integer managerId, Integer paymentTypeId, Integer deliveryTypeId, Integer statusId,
                /* Double priceDelivery,*/ Double priceFull,
                 /*Double payOther, */Double sumAllPays, String customerName, String mail, String phone,
                 String paymentType, String deliveryType, String deliveryAddress, String status, String manager,
                 String description, Date dateIncome, Date dateShipOut, Date dateStart, Date dateShipBack,
                 Date datePayment, Integer payInStatusId, Double payOut, Integer discount, Integer productShipStatusId,
                 String productShipStatusDescription, Date dateClose, Date salaryPayDate) {
        this.orderId = orderId;
        this.countProduct = countProduct;
        this.countDay = countDay;
        this.customerId = customerId;
        this.managerId = managerId;
        this.paymentTypeId = paymentTypeId;
        this.deliveryTypeId = deliveryTypeId;
        this.statusId = statusId;
//        this.priceDelivery = priceDelivery;
        this.priceFull = priceFull;
//        this.payOther = payOther;
        this.sumAllPays = sumAllPays;
        this.customerName = customerName;
        this.mail = mail;
        this.phone = phone;
        this.paymentType = paymentType;
        this.deliveryType = deliveryType;
        this.deliveryAddress = deliveryAddress;
        this.status = status;
        this.manager = manager;
        this.description = description;
        this.dateIncome = dateIncome;
        this.dateShipOut = dateShipOut;
        this.dateStart = dateStart;
        this.dateShipBack = dateShipBack;
        this.datePayment = datePayment;
        this.payInStatusId = payInStatusId;
        this.payOut = payOut;
        this.discount = discount;
        this.productShipStatusId = productShipStatusId;
        this.productShipStatusDescription = productShipStatusDescription;
        this.dateClose = dateClose;
        this.salaryPayDate = salaryPayDate;
    }

    public Integer getProfit() {
        profit = sumAllPays.intValue() - payOut.intValue();
        return profit;
    }
    public String getProfitString() {
        return getProfit().toString();
    }

    public Integer getOrderId() {
        return orderId;
    }

    public void setOrderId(Integer orderId) {
        this.orderId = orderId;
    }

    public Integer getCountProduct() {
        return countProduct;
    }

    public void setCountProduct(Integer countProduct) {
        this.countProduct = countProduct;
    }

    public Integer getCountDay() {
        return countDay;
    }

    public void setCountDay(Integer countDay) {
        this.countDay = countDay;
    }

    public Integer getCustomerId() {
        return customerId;
    }

    public Order setCustomerId(Integer customerId) {
        this.customerId = customerId;
        return this;
    }

    public Integer getManagerId() {
        return managerId;
    }

    public Order setManagerId(Integer managerId) {
        this.managerId = managerId;
        return this;
    }

    public Integer getPaymentTypeId() {
        return paymentTypeId;
    }

    public Order setPaymentTypeId(Integer paymentTypeId) {
        this.paymentTypeId = paymentTypeId;
        return this;
    }

    public Integer getDeliveryTypeId() {
        return deliveryTypeId;
    }

    public Order setDeliveryTypeId(Integer deliveryTypeId) {
        this.deliveryTypeId = deliveryTypeId;
        return this;
    }

    public Integer getStatusId() {
        return statusId;
    }

    public Integer getPayInStatusId() {
        return payInStatusId;
    }

    public void setPayInStatusId(Integer payInStatusId) {
        this.payInStatusId = payInStatusId;
    }

    public Order setStatusId(Integer statusId) {
        this.statusId = statusId;
        return this;
    }

//    public Double getPriceDelivery() {
//        return priceDelivery;
//    }
//
//    public void setPriceDelivery(Double priceDelivery) {
//        this.priceDelivery = priceDelivery;
//    }


    public Double getPayOut() {
        return payOut;
    }
    public String getPayOutString(){
        Integer sumOut = payOut != null ? payOut.intValue() : null;
        return sumOut != null ? sumOut.toString() : null;
    }

    public void setPayOut(Double payOut) {
        this.payOut = payOut;
    }

    public Double getPriceFull() {
        return priceFull;
    }

    public String getPriceFullString(){
        Integer price = priceFull != null ? priceFull.intValue() : null;
        return price != null ? price.toString() : null;
    }

    public void setPriceFull(Double priceFull) {
        this.priceFull = priceFull;
    }

//    public Double getPayOther() {
//        return payOther;
//    }
//
//    public void setPayOther(Double payOther) {
//        this.payOther = payOther;
//    }

    public Double getSumAllPays() {
        return sumAllPays;
    }

    public  String getSumAllPaysString(){
        Integer sumOut = sumAllPays != null ? sumAllPays.intValue() : null;
        return sumOut != null ? sumOut.toString() : null;
    }

    public void setSumAllPays(Double sumAllPays) {
        this.sumAllPays = sumAllPays;
    }

    public String getCustomerName() {
        return customerName;
    }

    public void setCustomerName(String customerName) {
        this.customerName = customerName;
    }

    public String getMail() {
        return mail;
    }

    public void setMail(String mail) {
        this.mail = mail;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getPaymentType() {
        return paymentType;
    }

    public void setPaymentType(String paymentType) {
        this.paymentType = paymentType;
    }

    public String getDeliveryType() {
        return deliveryType;
    }

    public void setDeliveryType(String deliveryType) {
        this.deliveryType = deliveryType;
    }

    public String getDeliveryAddress() {
        return deliveryAddress;
    }

    public void setDeliveryAddress(String deliveryAddress) {
        this.deliveryAddress = deliveryAddress;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public String getManager() {
        return manager;
    }

    public void setManager(String manager) {
        this.manager = manager;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public Date getDateIncome() {
        return dateIncome;
    }

    public String getDateIncomeString() {
        return dateIncome == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(dateIncome);
    }
    public String getDateIncomeStringDDMMYYYY() {
        return dateIncome == null ? "" : new SimpleDateFormat("dd/MM/yyyy").format(dateIncome);
    }

    public void setDateIncome(Date dateIncome) {
        this.dateIncome = dateIncome;
    }

    public Date getDateShipOut() {
        return dateShipOut;
    }

    public String getDateShipOutString() {
        return dateShipOut == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(dateShipOut);
    }
    public String getDateShipOutStringDDMMYYYY() {
        return dateShipOut == null ? "" : new SimpleDateFormat("dd/MM/yyyy").format(dateShipOut);
    }

    public void setDateShipOut(Date dateShipOut) {
        this.dateShipOut = dateShipOut;
    }

    public Date getDateStart() {
        return dateStart;
    }

    public String getDateStartString() {
        return dateStart == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(dateStart);
    }
    public String getDateStartStringDDMMYYYY() {
        return dateStart == null ? "" : new SimpleDateFormat("dd/MM/yyyy").format(dateStart);
    }

    public void setDateStart(Date dateStart) {
        this.dateStart = dateStart;
    }

    public Date getDateShipBack() {
        return dateShipBack;
    }

    public String getDateShipBackString() {
        return dateShipBack == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(dateShipBack);
    }
    public String getDateShipBackStringDDMMYYYY() {
        return dateShipBack == null ? "" : new SimpleDateFormat("dd/MM/yyyy").format(dateShipBack);
    }


    public void setDateShipBack(Date dateShipBack) {
        this.dateShipBack = dateShipBack;
    }

    public Date getDatePayment() {
        return datePayment;
    }

    public String getDatePaymentString() {
        return datePayment == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(datePayment);
    }

    public Date getDateClose() {
        return dateClose;
    }

    public String getDateCloseString() {
        return dateClose == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(dateClose);
    }

    public String getDateCloseStringDDMMYYYY() {
        return dateClose == null ? "" : new SimpleDateFormat("dd/MM/yyyy").format(dateClose);
    }

    public void setDateClose(Date dateClose) {
        this.dateClose = dateClose;
    }


    public Date getSalaryPayDate() {
        return salaryPayDate;
    }
    public String getSalaryPayDateString() {
        return salaryPayDate == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(salaryPayDate);
    }
    public String getSalaryPayDateStringDDMMYYYY() {
        return salaryPayDate == null ? "" : new SimpleDateFormat("dd/MM/yyyy").format(salaryPayDate);
    }

    public void setSalaryPayDate(Date salaryPayDate) {
        this.salaryPayDate = salaryPayDate;
    }

    public void setDatePayment(Date datePayment) {
        this.datePayment = datePayment;
    }

    public Integer getDiscount() {
        return discount;
    }

    public void setDiscount(Integer discount) {
        this.discount = discount;
    }

    public Integer getPriceDiscount(){
        priceFull = priceFull == null ? 0 : priceFull;
        discount = discount == null ? 0 : discount;
        Double price = priceFull  - ((priceFull / 100) * discount);
        Integer priceInt = price.intValue();
        return priceInt;
    }

    public String getPriceDiscountString(){
//        priceFull = priceFull == null ? 0 : priceFull;
//        discount = discount == null ? 0 : discount;
//        Double price = priceFull  - ((priceFull / 100) * discount);
//        Integer priceInt = price.intValue();
        return getPriceDiscount().toString();
    }

    public Integer getProductShipStatusId() {
        return productShipStatusId;
    }

    public void setProductShipStatusId(Integer productShipStatusId) {
        this.productShipStatusId = productShipStatusId;
    }

    public String getProductShipStatusDescription() {
        return productShipStatusDescription;
    }

    public void setProductShipStatusDescription(String productShipStatusDescription) {
        this.productShipStatusDescription = productShipStatusDescription;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Order order = (Order) o;
        return Objects.equals(orderId, order.orderId);
    }

    @Override
    public int hashCode() {
        return Objects.hash(orderId);
    }

    @Override
    public String toString() {
        return "Order{" +
                "orderId=" + orderId +
                ", countProduct=" + countProduct +
                ", countDay=" + countDay +
                ", customerId=" + customerId +
                ", managerId=" + managerId +
                ", paymentTypeId=" + paymentTypeId +
                ", deliveryTypeId=" + deliveryTypeId +
                ", statusId=" + statusId +
                ", payInStatusIs=" + payInStatusId +
//                ", priceDelivery=" + priceDelivery +
                ", priceFull=" + priceFull +
//                ", payOther=" + payOther +
                ", sumAllPays=" + sumAllPays +
                ", customerName='" + customerName + '\'' +
                ", mail='" + mail + '\'' +
                ", phone='" + phone + '\'' +
                ", paymentType='" + paymentType + '\'' +
                ", deliveryType='" + deliveryType + '\'' +
                ", deliveryAddress='" + deliveryAddress + '\'' +
                ", status='" + status + '\'' +
                ", manager='" + manager + '\'' +
                ", description='" + description + '\'' +
                ", dateIncome=" + dateIncome +
                ", dateShipOut=" + dateShipOut +
                ", dateStart=" + dateStart +
                ", dateShipBack=" + dateShipBack +
                ", datePayment=" + datePayment +
                ", dateClose=" + dateClose +
                ", salaryPayDate=" + salaryPayDate +
                ", payOut=" + payOut +
                ", discount=" + discount +
                ", productShipStatusId=" + productShipStatusId +
                ", productShipStatusDescription=" + productShipStatusDescription +
                '}';
    }
}
