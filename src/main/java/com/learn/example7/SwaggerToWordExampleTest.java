package com.learn.example7;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.Configure.ELMode;
import com.deepoove.poi.data.HyperLinkTextRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.BookmarkRenderPolicy;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import io.swagger.models.ArrayModel;
import io.swagger.models.HttpMethod;
import io.swagger.models.Model;
import io.swagger.models.ModelImpl;
import io.swagger.models.Path;
import io.swagger.models.RefModel;
import io.swagger.models.Swagger;
import io.swagger.models.parameters.AbstractSerializableParameter;
import io.swagger.models.parameters.BodyParameter;
import io.swagger.models.properties.AbstractNumericProperty;
import io.swagger.models.properties.ArrayProperty;
import io.swagger.models.properties.BooleanProperty;
import io.swagger.models.properties.Property;
import io.swagger.models.properties.RefProperty;
import io.swagger.parser.SwaggerParser;

/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description :
 */
@DisplayName("Swagger to Word Example")
public class SwaggerToWordExampleTest {

    String location = "D:\\test-poitl\\petstore.json";

    @Test
    public void testSwaggerToWord() throws IOException {
        SwaggerParser swaggerParser = new SwaggerParser();
        Swagger swagger = swaggerParser.read(location);
        SwaggerView viewData = convert(swagger);

        HackLoopTableRenderPolicy hackLoopTableRenderPolicy = new HackLoopTableRenderPolicy();
        Configure config = Configure.newBuilder().bind("parameters", hackLoopTableRenderPolicy).bind("responses", hackLoopTableRenderPolicy).bind("properties", hackLoopTableRenderPolicy).addPlugin('>', new BookmarkRenderPolicy()).setElMode(ELMode.SPEL_MODE).build();
        XWPFTemplate template = XWPFTemplate.compile("D:\\test-poitl\\swagger.docx", config).render(viewData);
        template.writeToFile("D:\\test-poitl\\swagger_out.docx");
    }

    private SwaggerView convert(Swagger swagger) {
        SwaggerView view = new SwaggerView();
        view.setInfo(swagger.getInfo());
        view.setBasePath(swagger.getBasePath());
        view.setExternalDocs(swagger.getExternalDocs());
        view.setHost(swagger.getHost());
        view.setSchemes(swagger.getSchemes());
        List<Resource> resources = new ArrayList<>();
        List<Endpoint> endpoints = new ArrayList<>();
        Map<String, Path> paths = swagger.getPaths();
        if (null == paths) {
            return view;
        }
        paths.forEach((url, path) -> {
            if (null == path.getOperationMap()) {
                return;
            }
            path.getOperationMap().forEach((method, operation) -> {
                Endpoint endpoint = new Endpoint();
                endpoint.setUrl(url);
                endpoint.setHttpMethod(method.toString());
                endpoint.setGet(HttpMethod.GET == method);
                endpoint.setDelete(HttpMethod.DELETE == method);
                endpoint.setPut(HttpMethod.PUT == method);
                endpoint.setPost(HttpMethod.POST == method);
                endpoint.setSummary(operation.getSummary());
                endpoint.setDescription(operation.getDescription());
                endpoint.setTag(operation.getTags());
                endpoint.setProduces(operation.getProduces());
                endpoint.setConsumes(operation.getConsumes());

                if (null != operation.getParameters()) {
                    List<Parameter> parameters = new ArrayList<>();
                    operation.getParameters().forEach(para -> {
                        Parameter parameter = new Parameter();
                        parameter.setDescription(para.getDescription());
                        parameter.setIn(para.getIn());
                        parameter.setName(para.getName());
                        parameter.setRequired(para.getRequired());
                        List<TextRenderData> schema = new ArrayList<>();
                        if (para instanceof AbstractSerializableParameter) {
                            Property items = ((AbstractSerializableParameter<?>) para).getItems();
                            String type = ((AbstractSerializableParameter<?>) para).getType();
                            // if array
                            if (ArrayProperty.isType(type)) {
                                schema.add(new TextRenderData("<"));
                                schema.addAll(formatProperty(items));
                                schema.add(new TextRenderData(">"));
                            } else {
                                schema.addAll(formatProperty(items));
                            }

                            if (StringUtils.isNotBlank(type)) {
                                schema.add(new TextRenderData(type));
                            }
                            if (StringUtils.isNotBlank(((AbstractSerializableParameter<?>) para).getCollectionFormat())) {
                                schema.add(new TextRenderData("(" + ((AbstractSerializableParameter<?>) para).getCollectionFormat() + ")"));
                            }
                        }
                        if (para instanceof BodyParameter) {
                            Model schemaModel = ((BodyParameter) para).getSchema();
                            schema.addAll(formatSchemaModel(schemaModel));
                        }
                        parameter.setSchema(schema);
                        parameters.add(parameter);
                    });
                    endpoint.setParameters(parameters);
                }

                if (null != operation.getResponses()) {
                    List<Response> responses = new ArrayList<>();
                    operation.getResponses().forEach((code, resp) -> {
                        Response response = new Response();
                        response.setCode(code);
                        response.setDescription(resp.getDescription());
                        Model schemaModel = resp.getResponseSchema();
                        response.setSchema(formatSchemaModel(schemaModel));

                        if (null != resp.getHeaders()) {
                            List<Header> headers = new ArrayList<>();
                            resp.getHeaders().forEach((name, property) -> {
                                Header header = new Header();
                                header.setName(name);
                                header.setDescription(property.getDescription());
                                header.setType(property.getType());
                                headers.add(header);
                            });
                            response.setHeaders(headers);
                        }
                        responses.add(response);
                    });
                    endpoint.setResponses(responses);
                }
                endpoints.add(endpoint);
            });
        });

        swagger.getTags().forEach(tag -> {
            Resource resource = new Resource();
            resource.setName(tag.getName());
            resource.setDescription(tag.getDescription());
            resource.setEndpoints(endpoints.stream().filter(path -> (null != path.getTag() && path.getTag().contains(tag.getName()))).collect(Collectors.toList()));
            resources.add(resource);
        });
        view.setResources(resources);

        ObjectMapper objectMapper = new ObjectMapper();
        if (null != swagger.getDefinitions()) {
            List<Definition> definitions = new ArrayList<>();
            swagger.getDefinitions().forEach((name, model) -> {
                Definition definition = new Definition();
                definition.setName(name);
                if (null != model.getProperties()) {
                    List<ExampleProperty> properties = new ArrayList<>();
                    model.getProperties().forEach((key, prop) -> {
                        ExampleProperty property = new ExampleProperty();
                        property.setName(key);
                        property.setDescription(prop.getDescription());
                        property.setRequired(prop.getRequired());
                        property.setSchema(formatProperty(prop));
                        properties.add(property);
                    });
                    definition.setProperties(properties);
                    Map<String, Object> map = valueOfModel(swagger.getDefinitions(), model, new HashSet<>());
                    try {
                        String writeValueAsString = objectMapper.writeValueAsString(map);
                        JsonNode readTree = objectMapper.readTree(writeValueAsString);
                        definition.setCodes(new JSONRenderPolicy("ffffff").convert(readTree, 1));
                    } catch (JsonProcessingException e) {
                        e.printStackTrace();
                    }
                }
                definitions.add(definition);
            });
            view.setDefinitions(definitions);
        }
        return view;
    }

    private Map<String, Object> valueOfModel(Map<String, Model> definitions, Model model, Set<String> keyCache) {
        Map<String, Object> map = new LinkedHashMap<>();
        model.getProperties().forEach((key, prop) -> {
            Object value = valueOfProperty(definitions, prop, keyCache);
            map.put(key, value);
        });
        return map;
    }

    private Object valueOfProperty(Map<String, Model> definitions, Property property, Set<String> keyCache) {
        Object value;
        if (property instanceof RefProperty) {
            String ref = ((RefProperty) property).get$ref().substring("#/definitions/".length());
            if (keyCache.contains(ref)) {
                value = ((RefProperty) property).get$ref();
            } else {
                value = valueOfModel(definitions, definitions.get(ref), keyCache);
            }
        } else if (property instanceof ArrayProperty) {
            List<Object> list = new ArrayList<>();
            Property insideItems = ((ArrayProperty) property).getItems();
            list.add(valueOfProperty(definitions, insideItems, keyCache));
            value = list;
        } else if (property instanceof AbstractNumericProperty) {
            value = 0;
        } else if (property instanceof BooleanProperty) {
            value = false;
        } else {
            value = property.getType();
        }
        return value;
    }

    private List<TextRenderData> formatSchemaModel(Model schemaModel) {
        List<TextRenderData> schema = new ArrayList<>();
        if (null == schemaModel) {
            return schema;
        }
        if (schemaModel instanceof ArrayModel) {
            Property items = ((ArrayModel) schemaModel).getItems();
            schema.add(new TextRenderData("<"));
            schema.addAll(formatProperty(items));
            schema.add(new TextRenderData(">"));
            schema.add(new TextRenderData(((ArrayModel) schemaModel).getType()));
        } else if (schemaModel instanceof RefModel) {
            String ref = ((RefModel) schemaModel).get$ref().substring("#/definitions/".length());
            schema.add(new HyperLinkTextRenderData(ref, "anchor:" + ref));
        } else if (schemaModel instanceof ModelImpl) {
            schema.add(new TextRenderData(((ModelImpl) schemaModel).getType()));
        }
        return schema;
    }

    private List<TextRenderData> formatProperty(Property items) {
        List<TextRenderData> schema = new ArrayList<>();
        if (null != items) {
            if (items instanceof RefProperty) {
                String ref = ((RefProperty) items).get$ref().substring("#/definitions/".length());
                schema.add(new HyperLinkTextRenderData(ref, "anchor:" + ref));
            } else if (items instanceof ArrayProperty) {
                Property insideItems = ((ArrayProperty) items).getItems();
                schema.add(new TextRenderData("<"));
                schema.addAll(formatProperty(insideItems));
                schema.add(new TextRenderData(">"));
                schema.add(new TextRenderData(items.getType()));
            } else {
                schema.add(new TextRenderData(items.getType()));
            }
        }
        return schema;
    }
}