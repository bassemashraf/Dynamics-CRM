export interface Step {
    duc_sequence: number
    duc_sequenceoverride: number | null
    duc_visible: boolean | null
    duc_servicerequesttypestepsid: string // guid
    duc_name: string | null
    duc_arabicname: string | null
    duc_descriptionen: string | null
    duc_description: string | null
    IsTempStep: boolean | null
}


export interface ActionLog {
    duc_ServiceRequestTypeStep: Step
}