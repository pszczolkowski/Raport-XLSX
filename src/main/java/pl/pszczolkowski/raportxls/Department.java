package pl.pszczolkowski.raportxls;

public class Department {

    private final String name;
    private final String ownerName;

    public Department(String name, String ownerName) {
        this.name = name;
        this.ownerName = ownerName;
    }

    public String getName() {
        return name;
    }

    public String getOwnerName() {
        return ownerName;
    }

}
