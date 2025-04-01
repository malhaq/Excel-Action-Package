package com.automationanywhere.botcommand;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class DeleteByConditionCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(DeleteByConditionCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    DeleteByCondition command = new DeleteByCondition();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("inputFilePath") && parameters.get("inputFilePath") != null && parameters.get("inputFilePath").get() != null) {
      convertedParameters.put("inputFilePath", parameters.get("inputFilePath").get());
      if(convertedParameters.get("inputFilePath") !=null && !(convertedParameters.get("inputFilePath") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","inputFilePath", "String", parameters.get("inputFilePath").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("inputFilePath") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","inputFilePath"));
    }

    if(parameters.containsKey("outputFilePath") && parameters.get("outputFilePath") != null && parameters.get("outputFilePath").get() != null) {
      convertedParameters.put("outputFilePath", parameters.get("outputFilePath").get());
      if(convertedParameters.get("outputFilePath") !=null && !(convertedParameters.get("outputFilePath") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","outputFilePath", "String", parameters.get("outputFilePath").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("outputFilePath") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","outputFilePath"));
    }

    if(parameters.containsKey("column") && parameters.get("column") != null && parameters.get("column").get() != null) {
      convertedParameters.put("column", parameters.get("column").get());
      if(convertedParameters.get("column") !=null && !(convertedParameters.get("column") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","column", "String", parameters.get("column").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("column") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","column"));
    }

    if(parameters.containsKey("Condition") && parameters.get("Condition") != null && parameters.get("Condition").get() != null) {
      convertedParameters.put("Condition", parameters.get("Condition").get());
      if(convertedParameters.get("Condition") !=null && !(convertedParameters.get("Condition") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","Condition", "String", parameters.get("Condition").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("Condition") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","Condition"));
    }
    if(convertedParameters.get("Condition") != null) {
      switch((String)convertedParameters.get("Condition")) {
        case "Equals to =" : {

        } break;
        case "Doesn't equals !=" : {

        } break;
        case "Contain" : {

        } break;
        case "Doesn't Contain" : {

        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","Condition"));
      }
    }

    if(parameters.containsKey("criteriaValue") && parameters.get("criteriaValue") != null && parameters.get("criteriaValue").get() != null) {
      convertedParameters.put("criteriaValue", parameters.get("criteriaValue").get());
      if(convertedParameters.get("criteriaValue") !=null && !(convertedParameters.get("criteriaValue") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","criteriaValue", "String", parameters.get("criteriaValue").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("criteriaValue") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","criteriaValue"));
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("inputFilePath"),(String)convertedParameters.get("outputFilePath"),(String)convertedParameters.get("column"),(String)convertedParameters.get("Condition"),(String)convertedParameters.get("criteriaValue")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }

  public Map<String, Value> executeAndReturnMany(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return null;
  }
}
