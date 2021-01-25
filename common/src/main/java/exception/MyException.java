package exception;

public class MyException {

    public static <T> T notNull(T object, String errorMsg) throws IllegalArgumentException {
        if (object == null) {
            throw new IllegalArgumentException(errorMsg);
        }
        return object;
    }
}
