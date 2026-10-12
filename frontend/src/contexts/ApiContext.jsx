import React, { createContext, useContext } from "react";
const ApiContext = createContext({ baseUrl: import.meta.env.VITE_API_BASE_URL || "http://127.0.0.1:8000" });
export const ApiProvider = ({ children }) => <ApiContext.Provider value={{ baseUrl: import.meta.env.VITE_API_BASE_URL || "http://127.0.0.1:8000" }}>{children}</ApiContext.Provider>;
export const useApi = () => useContext(ApiContext);
