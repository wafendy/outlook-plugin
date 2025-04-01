import { api } from "../api/http.client";
import { Version } from "../interfaces/version.interface";

class VersionService {
  public getVersion = async (): Promise<Version> => {
    return api.get("/version").then((resp) => resp.data);
  };
}

export const versionService = new VersionService();
