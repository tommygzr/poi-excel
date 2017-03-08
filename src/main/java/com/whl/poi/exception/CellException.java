package com.whl.poi.exception;

public class CellException extends  RuntimeException {

	private static final long serialVersionUID = -8014001549642815373L;
	private String code;

	public CellException() {
		this.code = null;
	}

	public CellException(final String message) {
		super(message);
		this.code = null;
	}

	public CellException(final Throwable cause) {
		super(cause);
		this.code = null;
	}

	public CellException(final String message, final Throwable cause) {
		super(message, cause);
		this.code = null;
	}

	public CellException(final String code, final String message) {
		super(message);
		this.code = code;
	}

	public CellException(final String code, final String message, final Throwable cause) {
		super(message, cause);
		this.code = code;
	}

	public String getCode() {
		return this.code;
	}
}
