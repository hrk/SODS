package com.github.miachm.sods;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.util.HashMap;
import java.util.Map;

class XmlReaderInstanceEventImpl implements XmlReaderInstance {
    private XMLStreamReader reader;
    private String tag;
    private String characters;
    private boolean end = false;
    private Map<String, String> atributes = new HashMap<>();

    XmlReaderInstanceEventImpl(XMLStreamReader reader, String tag)
    {
        this.reader = reader;
        this.tag = tag;

        if (reader.isStartElement())
            for (int i = 0; i < reader.getAttributeCount(); i++) {
                String name = qNameToString(reader.getAttributeName(i));
                String value = reader.getAttributeValue(i);
                atributes.put(name, value);
            }
        else if (reader.isCharacters()) {
            characters = reader.getText();
            end = true;
        }
    }

    @Override
    public boolean hasNext() {
        try {
            return reader.hasNext() && !end;
        } catch (XMLStreamException e) {
            throw new NotAnOdsException(e);
        }
    }

    @Override
    public XmlReaderInstance nextElement(String... names) {
        try {
            while (reader.hasNext() && !end) {
                reader.next();
                if (reader.isStartElement()) {
                    QName qName = reader.getName();
                    String elementName = qNameToString(qName);
                    if (contains(names, elementName))
                        return new XmlReaderInstanceEventImpl(reader, elementName);
                }
                else if (reader.isEndElement()) {
                    QName qName = reader.getName();
                    String elementName = qNameToString(qName);
                    if (elementName.equals(tag)) {
                        end = true;
                        return null;
                    }
                }
                else if (reader.isCharacters()) {
                    if (contains(names, CHARACTERS))
                        return new XmlReaderInstanceEventImpl(reader, CHARACTERS);
                }
            }
            return null;
        }
        catch(XMLStreamException e){
            throw new NotAnOdsException(e);
        }
    }

    @Override
    public String getAttribValue(String name) {
        return atributes.get(name);
    }

    @Override
    public Map<String, String> getAllAttributes()
    {
        return atributes;
    }

    @Override
    public String getContent()
    {
        return characters;
    }

    @Override
    public String getTag() {
        return tag;
    }

    private String qNameToString(QName qName)
    {
        return qName.getPrefix() + ":" + qName.getLocalPart();
    }

    public static <String> boolean contains(final String[] array, final String v) {
        for (final String e : array)
            if (e.equals(v))
                return true;

        return false;
    }
}
