package io.github.prajyotsable.exception.handler;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.multipart.MaxUploadSizeExceededException;

import java.util.Map;

@RestControllerAdvice
public class GlobalExceptionHandler {

    @ExceptionHandler(MaxUploadSizeExceededException.class)
    public ResponseEntity<Map<String, String>> handleMaxSizeException(
            MaxUploadSizeExceededException ex) {
        
        return ResponseEntity
                .status(HttpStatus.PAYLOAD_TOO_LARGE)  // 413
                .body(Map.of(
                    "error", "File too large",
                    "message", "Maximum allowed size is 200MB"
                ));
    }
}