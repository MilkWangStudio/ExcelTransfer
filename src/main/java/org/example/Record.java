package org.example;

import lombok.Data;

@Data
public class Record {
    private String type;
    private String owner;
    private String question;
    private String resolver;
    private String answer;
}
