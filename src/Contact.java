import java.util.Objects;

public class Contact {

    private Integer contactId;
    private Integer employeeId;
    private Integer customerId;
    private Integer contactTypeId;
    private String userName;
    private String contactType;
    private String contactData;
    private String flagMainContact;

    public Contact() {
    }

    public Contact(Integer contactId, Integer employeeId, Integer customerId, Integer contactTypeId,
                   String userName, String contactType, String contactData, String flagMainContact) {
        this.contactId = contactId;
        this.employeeId = employeeId;
        this.customerId = customerId;
        this.contactTypeId = contactTypeId;
        this.userName = userName;
        this.contactType = contactType;
        this.contactData = contactData;
        this.flagMainContact = flagMainContact;
    }

    public Integer getContactId() {
        return contactId;
    }

    public void setContactId(Integer contactId) {
        this.contactId = contactId;
    }

    public Integer getEmployeeId() {
        return employeeId;
    }

    public void setEmployeeId(Integer employeeId) {
        this.employeeId = employeeId;
    }

    public Integer getCustomerId() {
        return customerId;
    }

    public void setCustomerId(Integer customerId) {
        this.customerId = customerId;
    }

    public Integer getContactTypeId() {
        return contactTypeId;
    }

    public void setContactTypeId(Integer contactTypeId) {
        this.contactTypeId = contactTypeId;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getContactType() {
        return contactType;
    }

    public void setContactType(String contactType) {
        this.contactType = contactType;
    }

    public String getContactData() {
        return contactData;
    }

    public void setContactData(String contactData) {
        this.contactData = contactData;
    }

    public String getFlagMainContact() {
        return flagMainContact;
    }

    public void setFlagMainContact(String flagMainContact) {
        this.flagMainContact = flagMainContact;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Contact contact = (Contact) o;
        return Objects.equals(contactId, contact.contactId);
    }

    @Override
    public int hashCode() {
        return Objects.hash(contactId);
    }

    @Override
    public String toString() {
        return "Contact{" +
                "contactId=" + contactId +
                ", employeeId=" + employeeId +
                ", customerId=" + customerId +
                ", contactTypeId=" + contactTypeId +
                ", userName='" + userName + '\'' +
                ", contactType='" + contactType + '\'' +
                ", contactData='" + contactData + '\'' +
                ", flagMainContact='" + flagMainContact + '\'' +
                '}';
    }
}
