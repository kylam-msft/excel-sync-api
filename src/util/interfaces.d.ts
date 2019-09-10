interface IWorkerLogMessagePayload {
    onBehalfOf: 'client' | 'worker' | 'system';
    data: any;
    severity: ConsoleLogTypes;
}

interface IWorkerDonePayload {
    numberOfSyncRoundtrips: number;
    workerTotalElapsedTime: number;
    timeSpentWaiting: number;
}

type ConsoleLogTypes = "log" | "info" | "warn" | "error";
