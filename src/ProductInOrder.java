import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;

public class ProductInOrder {


    private Integer orderAndProductId;
    private Integer orderId;
    private Integer productId;
    private String productName;
    private String productModel;
    private Integer productAmount;
    private Double productPrice;
    private Double totalPrice;
    private Date shipOutDate;
    private Date shipBackDate;

    public ProductInOrder() {
    }

    public ProductInOrder(Integer orderAndProductId, Integer orderId, Integer productId,
                          String productName, String productModel, Integer productAmount,
                          Double productPrice, Double totalPrice, Date shipOutDate, Date shipBackDate) {
        this.orderAndProductId = orderAndProductId;
        this.orderId = orderId;
        this.productId = productId;
        this.productName = productName;
        this.productModel = productModel;
        this.productAmount = productAmount;
        this.productPrice = productPrice;
        this.totalPrice = totalPrice;
        this.shipOutDate = shipOutDate;
        this.shipBackDate = shipBackDate;
    }

    public Integer getOrderAndProductId() {
        return orderAndProductId;
    }

    public void setOrderAndProductId(Integer orderAndProductId) {
        this.orderAndProductId = orderAndProductId;
    }

    public Integer getOrderId() {
        return orderId;
    }

    public void setOrderId(Integer orderId) {
        this.orderId = orderId;
    }

    public Integer getProductId() {
        return productId;
    }

    public void setProductId(Integer productId) {
        this.productId = productId;
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName(String productName) {
        this.productName = productName;
    }

    public String getProductModel() {
        return productModel;
    }

    public void setProductModel(String productModel) {
        this.productModel = productModel;
    }

    public Integer getProductAmount() {
        return productAmount;
    }
    public String getProductAmountString() {
        return productAmount != null ? productAmount.toString() : null;
    }

    public void setProductAmount(Integer productAmount) {
        this.productAmount = productAmount;
    }

    public Double getProductPrice() {
        return productPrice;
    }
    public String getProductPriceString() {
        return productPrice != null ? productPrice.toString() : null;
    }

    public void setProductPrice(Double productPrice) {
        this.productPrice = productPrice;
    }

    public Double getTotalPrice() {
        return totalPrice;
    }

    public void setTotalPrice(Double totalPrice) {
        this.totalPrice = totalPrice;
    }

    public Date getShipOutDate() {
        return shipOutDate;
    }
    public String getShipOutDateString () { return shipOutDate == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(shipOutDate); }
    public String getShipOutDateStringDDMMYYYY () { return shipOutDate == null ? "" : new SimpleDateFormat("dd/MM/yyyy").format(shipOutDate); }

    public void setShipOutDate(Date shipOutDate) {
        this.shipOutDate = shipOutDate;
    }

    public Date getShipBackDate() {
        return shipBackDate;
    }
    public String getShipBackDateeString () { return shipBackDate == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(shipBackDate); }
    public String getShipBackDateStringDDMMYYYY () { return shipBackDate == null ? "" : new SimpleDateFormat("dd/MM/yyyy").format(shipBackDate); }

    public void setShipBackDate(Date shipBackDate) {
        this.shipBackDate = shipBackDate;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        ProductInOrder that = (ProductInOrder) o;
        return Objects.equals(orderAndProductId, that.orderAndProductId);
    }

    @Override
    public int hashCode() {
        return Objects.hash(orderAndProductId);
    }

    @Override
    public String toString() {
        return "ProductInOrder{" +
                "orderAndProductId=" + orderAndProductId +
                ", orderId=" + orderId +
                ", productId=" + productId +
                ", productName='" + productName + '\'' +
                ", productModel='" + productModel + '\'' +
                ", productAmount=" + productAmount +
                ", productPrice=" + productPrice +
                ", totalPrice=" + totalPrice +
                ", shipOutDate=" + shipOutDate +
                ", shipBackDate=" + shipBackDate +
                '}';
    }
}
