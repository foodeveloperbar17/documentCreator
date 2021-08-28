package ge.luka;

import java.util.List;

public class DocumentModel {

    private String driverFullName;
    private String clinicName;
    private List<Client> clients;
    private String startHour;
    private String endHour;
    private String day;
    private String carId;
    private String blockId;

    public String getDriverFullName() {
        return driverFullName;
    }

    public void setDriverFullName(String driverFullName) {
        this.driverFullName = driverFullName;
    }

    public List<Client> getClients() {
        return clients;
    }

    public void setClients(List<Client> clients) {
        this.clients = clients;
    }

    public String getStartHour() {
        return startHour;
    }

    public void setStartHour(String startHour) {
        this.startHour = startHour;
    }

    public String getEndHour() {
        return endHour;
    }

    public void setEndHour(String endHour) {
        this.endHour = endHour;
    }

    public String getClinicName() {
        return clinicName;
    }

    public void setClinicName(String clinicName) {
        this.clinicName = clinicName;
    }

    public String getDay() {
        return day;
    }

    public void setDay(String day) {
        this.day = day;
    }

    public String getCarId() {
        return carId;
    }

    public void setCarId(String carId) {
        this.carId = carId;
    }

    public String getBlockId() {
        return blockId;
    }

    public void setBlockId(String blockId) {
        this.blockId = blockId;
    }

    @Override
    public String toString() {
        return "DocumentModel{" +
                "driverFullName='" + driverFullName + '\'' +
                ", clinicName='" + clinicName + '\'' +
                ", clients=" + clients +
                ", startHour='" + startHour + '\'' +
                ", endHour='" + endHour + '\'' +
                ", day='" + day + '\'' +
                ", carId='" + carId + '\'' +
                ", blockId='" + blockId + '\'' +
                '}';
    }
}
