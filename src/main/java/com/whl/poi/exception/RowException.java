package com.whl.poi.exception;

public class RowException extends  RuntimeException {

	private static final long serialVersionUID = -8014001549642815373L;
	private String code;

	public RowException() {
		this.code = null;
	}

	public RowException(final String message) {
		super(message);
		this.code = null;
	}

	public RowException(final Throwable cause) {
		super(cause);
		this.code = null;
	}

	public RowException(final String message, final Throwable cause) {
		super(message, cause);
		this.code = null;
	}

	public RowException(final String code, final String message) {
		super(message);
		this.code = code;
	}

	public RowException(final String code, final String message, final Throwable cause) {
		super(message, cause);
		this.code = code;
	}

	public String getCode() {
		return this.code;
	}
}
