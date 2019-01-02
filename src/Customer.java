import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Objects;

public class Customer {

    private Integer customerId;
    private Integer customerTypeId;
    private String customerTypeDescription;
    private String fName;
    private String sName;
    private String pName;
    private String dopInfo;
    private Date dateCreate;
    private List<Contact> mainPhones;
    private List<Contact> otherPhones;
    private List<Contact> mainMails;
    private List<Contact> otherMails;
    private List<Contact> vibers;
    private List<Contact> telegrams;
    private List<Contact> skypes;
    private Integer discount;

    public Customer() {
    }

    public Customer(Integer customerId, Integer customerTypeId, String customerTypeDescription,
                    String fName, String sName, String pName,
                    String dopInfo, Date dateCreate, Integer discount) {
        this.customerId = customerId;
        this.customerTypeId = customerTypeId;
        this.customerTypeDescription = customerTypeDescription;
        this.fName = fName;
        this.sName = sName;
        this.pName = pName;
        this.dopInfo = dopInfo;
        this.dateCreate = dateCreate;
        this.discount = discount;
    }

    public Customer(Integer customerId, Integer customerTypeId, String customerTypeDescription,
                    String fName, String sName, String pName,
                    String dopInfo, Date dateCreate, Integer discount,
                    List<Contact> mainPhones, List<Contact> otherPhones,
                    List<Contact> mainMails, List<Contact> otherMails, List<Contact> vibers,
                    List<Contact> telegrams, List<Contact> skypes) {
        this.customerId = customerId;
        this.customerTypeId = customerTypeId;
        this.customerTypeDescription = customerTypeDescription;
        this.fName = fName;
        this.sName = sName;
        this.pName = pName;
        this.dopInfo = dopInfo;
        this.dateCreate = dateCreate;
        this.discount = discount;
        this.mainPhones = mainPhones;
        this.otherPhones = otherPhones;
        this.mainMails = mainMails;
        this.otherMails = otherMails;
        this.vibers = vibers;
        this.telegrams = telegrams;
        this.skypes = skypes;
    }

    public Integer getCustomerId() {
        return customerId;
    }

    public void setCustomerId(Integer customerId) {
        this.customerId = customerId;
    }

    public Integer getCustomerTypeId() {
        return customerTypeId;
    }

    public Integer getDiscount() {
        return discount;
    }

    public void setDiscount(Integer discount) {
        this.discount = discount;
    }

    public void setCustomerTypeId(Integer customerTypeId) {
        this.customerTypeId = customerTypeId;
    }

    public String getCustomerTypeDescription() {
        return customerTypeDescription;
    }

    public void setCustomerTypeDescription(String customerTypeDescription) {
        this.customerTypeDescription = customerTypeDescription;
    }

    public String getfName() {
        return fName;
    }

    public void setfName(String fName) {
        this.fName = fName;
    }

    public String getsName() {
        return sName;
    }

    public void setsName(String sName) {
        this.sName = sName;
    }

    public String getpName() {
        return pName;
    }

    public void setpName(String pName) {
        this.pName = pName;
    }

    public String getFIO(){
        return (fName != null ? fName + " " : "") +(sName != null ? sName + " " : "") + (pName != null ? pName : "");
    }

    public String getDopInfo() {
        return dopInfo;
    }

    public void setDopInfo(String dopInfo) {
        this.dopInfo = dopInfo;
    }

    public Date getDateCreate() {
        return dateCreate;
    }

    public String getDateCreateString() {
        return dateCreate == null ? "" : new SimpleDateFormat("yyyy-MM-dd").format(dateCreate);
    }

    public void setDateCreate(Date dateCreate) {
        this.dateCreate = dateCreate;
    }

    public List<Contact> getMainPhones() {
        return mainPhones;
    }

    public void setMainPhones(List<Contact> mainPhones) {
        this.mainPhones = mainPhones;
    }

    public List<Contact> getOtherPhones() {
        return otherPhones;
    }

    public void setOtherPhones(List<Contact> otherPhones) {
        this.otherPhones = otherPhones;
    }

    public List<Contact> getMainMails() {
        return mainMails;
    }

    public void setMainMails(List<Contact> mainMails) {
        this.mainMails = mainMails;
    }

    public List<Contact> getOtherMails() {
        return otherMails;
    }

    public void setOtherMails(List<Contact> otherMails) {
        this.otherMails = otherMails;
    }

    public List<Contact> getVibers() {
        return vibers;
    }

    public void setVibers(List<Contact> vibers) {
        this.vibers = vibers;
    }

    public List<Contact> getTelegrams() {
        return telegrams;
    }

    public void setTelegrams(List<Contact> telegrams) {
        this.telegrams = telegrams;
    }

    public List<Contact> getSkypes() {
        return skypes;
    }

    public void setSkypes(List<Contact> skypes) {
        this.skypes = skypes;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Customer customer = (Customer) o;
        return Objects.equals(customerId, customer.customerId);
    }

    @Override
    public int hashCode() {
        return Objects.hash(customerId);
    }

    @Override
    public String toString() {
        return "Customer{" +
                "customerId=" + customerId +
                ", customerTypeId=" + customerTypeId +
                ", customerTypeDescription=" + customerTypeDescription +
                ", fName='" + fName + '\'' +
                ", sName='" + sName + '\'' +
                ", pName='" + pName + '\'' +
                ", dopInfo='" + dopInfo + '\'' +
                ", dateCreate=" + dateCreate +
                ", discount=" + discount +
                ", mainPhones=" + mainPhones +
                ", otherPhones=" + otherPhones +
                ", mainMails=" + mainMails +
                ", otherMails=" + otherMails +
                ", vibers=" + vibers +
                ", telegrams=" + telegrams +
                ", skypes=" + skypes +
                '}';
    }
}