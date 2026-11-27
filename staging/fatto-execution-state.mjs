// Pure state machine, not a durable executor. External effects require reconciliation.
export function startExecution(plan){
 if(!plan?.steps?.length||plan.steps[0].action!=='saveCompletion')throw Error('Invalid plan');
 return Object.freeze({plan,steps:Object.freeze(plan.steps.map(s=>Object.freeze({...s,status:'pending'})))});
}
export function transitionExecution(state,key,event){
 const index=state.steps.findIndex(s=>s.key===key);
 if(index<0)throw Error('Unknown step');
 const step=state.steps[index];
 const transitions={pending:{start:'running'},running:{success:'done',failure:'failed',uncertain:'uncertain'},failed:{retry:'pending'},uncertain:{confirmed:'done',notExecuted:'pending'},done:{}};
 const status=transitions[step.status]?.[event];
 if(!status||!Object.hasOwn(transitions[step.status],event))throw Error('Invalid transition');
 if(event==='start'&&state.steps.slice(0,index).some(s=>s.status!=='done'))throw Error('Previous step incomplete');
 return Object.freeze({...state,steps:Object.freeze(state.steps.map((s,i)=>i===index?Object.freeze({...s,status}):s))});
}
