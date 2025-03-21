import { useQuery } from "@tanstack/react-query";
import { versionService } from "../../services/version.service";

export const useGetVersionQuery = () => {
  return useQuery({
    staleTime: 5000,
    queryKey: ["version"],
    queryFn: async () => {
      return versionService.getVersion();
    },
  });
};
