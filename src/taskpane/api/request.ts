import type { AxiosRequestConfig } from "axios";
import axios from "axios";

const defaultConfig = {
  headers: {
    Accept: "application/json",
    "Content-Type": "application/json",
  },
};
const instance = axios.create(defaultConfig);
instance.defaults.timeout = 30000;

interface RequestParams extends AxiosRequestConfig {
  cache?: boolean;
  skipVersionCheck?: boolean;
}

// 未包装业务代码逻辑的请求
export const requestPure = async (params: RequestParams) => {
  const response = await instance(params);

  return response.data;
};
