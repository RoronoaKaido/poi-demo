import jdk.nashorn.internal.objects.annotations.Getter;

import java.util.List;
public class Tag {
    private Long id;

    private String name;

    private List<Tag> chilrenList;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Tag> getChilrenList() {
        return chilrenList;
    }

    public void setChilrenList(List<Tag> chilrenList) {
        this.chilrenList = chilrenList;
    }

    public Tag(Long id, String name, List<Tag> chilrenList) {
        this.id = id;
        this.name = name;
        this.chilrenList = chilrenList;
    }

    public Tag(Long id, String name) {
        this.id = id;
        this.name = name;
    }
}
