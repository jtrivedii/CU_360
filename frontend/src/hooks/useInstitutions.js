import { useEffect, useState } from "react";
import { institutionService } from "../services/api";
export function useInstitutions() { const [data,setData]=useState(null); const [error,setError]=useState(null); useEffect(()=>{let live=true; institutionService.list().then(v=>live&&setData(v)).catch(e=>live&&setError(e)); return ()=>{live=false};},[]); return {data,error,loading:!data&&!error}; }
