import { MSGraphClient } from "@microsoft/sp-http";
import { IPlannerBucket } from "../models/IPlannerBucket";
import { IPlannerPlan } from "../models/IPlannerPlan";
import { IPlannerTask } from "../models/IPlannerTask";
import { IUser } from "../models/IUser";

export async function getGroupMembers(groupGuid: string, client: MSGraphClient): Promise<IUser[]> {
    const users: IUser[] = [];
    const response = await client.api(`/groups/${groupGuid}/members?$expand=*`).get();
    if (response && response.value) {
        response.value.forEach(element => {
            const picture = "";

            users.push({
                id: element.id,
                displayName: element.displayName,
                userPrincipalName: element.userPrincipalName,
                mail: element.mail,
                profilePicture: picture
            });
        });
    }
    return users;
}

export async function getPlannerPlans(groupGuid: string, client: MSGraphClient): Promise<IPlannerPlan[]> {
    const plans: IPlannerPlan[] = [];
    const response = await client.api(`/groups/${groupGuid}/planner/plans`).get();
    if (response && response.value) {
        response.value.forEach(element => {
            const planTitle = element.title;
            const planGuid = element.id;
            plans.push({
                groupId: groupGuid,
                id: planGuid,
                title: planTitle
            });
        });
    }
    return plans;
}

export async function getPlannerBuckets(planGuid: string, client: MSGraphClient): Promise<IPlannerBucket[]> {
    const response = await client.api(`/planner/plans/${planGuid}/buckets`).get();

    const buckets: IPlannerBucket[] = [];
    if (response && response.value) {
        buckets.push({
            id: "",
            title: "Show all buckets",
            planId: planGuid
        });
        response.value.forEach(element => {
            const bucketTitle = element.name;
            const bucketGuid = element.id;
            buckets.push({
                id: bucketGuid,
                title: bucketTitle,
                planId: planGuid
            });
        });
    }
    return buckets;
}

export async function getPlannerTasks(planGuid: string, users: IUser[], client: MSGraphClient): Promise<IPlannerTask[]> {
    const response = await client.api(`/planner/plans/${planGuid}/tasks?$orderby=dueDateTime`).get();

    let tasks: IPlannerTask[] = [];
    if (response && response.value) {
        response.value.forEach(element => {
            const bucketGuid = element.bucketId;
            const taskId = element.id;
            const taskTitle = element.title;

            const dueDateTime = (element.dueDateTime === null ? new Date(2000, 1, 1) : element.dueDateTime);
            const startDateTime = (element.startDateTime === null && dueDateTime !== null ? dueDateTime : element.startDateTime);

            const completionState = element.percentComplete;
            const assignments: IUser[] = [];
            if (element.assignments !== undefined) {
                users.forEach(groupMember => {
                    try {
                        const assignment = element.assignments["" + groupMember.id];
                        if (assignment !== undefined && assignment !== null) {
                            assignments.push(groupMember);
                        }
                    }
                    catch (ex) { 
                        console.log(ex);
                    }
                });
            }
            tasks.push({
                id: taskId,
                title: taskTitle,
                dueDate: dueDateTime,
                startDate: startDateTime,
                percentComplete: completionState,
                assignedTo: assignments,
                bucketId: bucketGuid,
                planId: planGuid
            });
        });

        const sortedTasks = tasks?.sort((taskA,taskB) =>{
            if(taskA.startDate > taskB.startDate)
                return 1;
            else if (taskB.startDate > taskA.startDate)
                return -1
            else
                return 0;
        });
        tasks = sortedTasks;
    }
    return tasks;
}
