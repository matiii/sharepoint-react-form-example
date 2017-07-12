export interface BaseResponse { }

export interface ErrorResponse extends BaseResponse {
    message: string;
}

export interface SuccessResponse<T> extends BaseResponse {
    data: T;
}