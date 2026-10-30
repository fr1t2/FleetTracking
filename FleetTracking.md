# FleetTracking

## Overall Design

This dashboard has 2 primary functions. Viewing the current deployment and deploying technicians in vehicles. We have over 100 techs that we deploy on various schedules through the day. We need to track these vehicles by their status, and show who is driving and what zone they are in.

### Deployment

During deployment we need to be able to assign a vehicle, and a zone, to a tech, give them a task and have it shown deployed into that zone with that task assigned to them. 

We need to be able to re-assign or modify all of these details through the day. We should track and display the changes that happened for historical purpose. 

There are times when we double up techs in a vehicle, and at times these are also training sessions. These need to be shown differently so we can see them standing out. 

We also may send a tech to multiple zones and this should be shown on their assignment card. 

### Historical View

We need the ability to lookup past shifts and see the deployment in this dashboard. I envision a dropdown that lets me pick a day and one for the shift (1, 2, or 3) and pulls up the view that was for the past shift. These details should not be editable now as the shift is completed and this is the record. 

This all should be in the same view and allow export to CSV and PDF with a button that allows me to copy to clipboard the CSV of tech, zone and van assignments that are deployed. 


### Vehicles 

When a vehicle has an issue or breaks down, we need a way to track the issue, and depending on how severe, mark the vehicle out of service so it is not re-deployed.

Each vehicle has attributes and information that follows the van, including license plate #, VIN etc. We need to work on refining the issue tracker on the vans in this re-work.

When we add a new vehicle to the fleet we need to collect a bunch of information and once saved, prompt with a popup to also enter the information into the Samsara system. This popup should have instruction to instruct the user what needs to be done.

### Issue Tracking

Vehicles have issues that come up through their use, we need a way to track these, enter the issue and who reported/caused it, have an issue number assigned and show through an icon that there is service needed for that vehicle. The list of on-going issues should pop up when we deploy the vehicle so we can discuss any limitations with the driver. These issues could be low tire pressure, oil change notice, Broken lights or mirrors or doors not working. We should leave room for custom entry here.

If the van is not safe to drive due to the issue we need to mark it out of service and get a repair scheduled. When it is out of service we should show it in a different way to identify it is not to be dispatched. 

When service is completed on the vehicle we need to be able to mark the service complete, leave a note on what was done and track when and who completed the service. We can have a way to lookup historical service records but the old completed records do not need to show at dispatch or on the card. 

### Shift info

We run 3 shifts, first, second and third. First starts at 5am and runs waves of dispatch from 5-7 every half hour. So we have techs coming in at 5, 5:30, 6, 6:30 and 7. They work 8 hours and return to the warehouse to turn in their van for the next shift to dispatch. Next shift starts their waves from 1pm till 3pm and night does theirs from 9 - 11 also every half hour.

Vans come back based on when they are sent out. So the vans sent at 5am will be the first ones back at 1pm. We should have a timer run that gives an estimate on when a van will be back next and change an icon or color of the vehicle card to indicate what is expected to be back next. 

We should also display what shift the tech works prominently so we can discern if they are ending (heading back) or just deployed. So for example the van/shift card could have a green highlight when freshly deployed and shows a 1 in a corner signifying they are on the first shift. Since they deployed at 5 they will be back at 1. At around 12:30 we could change their highlight color to yellow so second shift can see they are soon returning and they work the 1st shift.

## Process

1. Start with a review of the current code and evaluate any changes that need to be made to accommodate this update
2. Lets layout a plan together including what needs to be changed, develop a scope document that you will follow and agree on the next step. You will need to bring any scoping questions and clarifications to me here so we are clear and understand the constraints and limitations as we move forward.
3. We will work through this plan and review as we go to make sure we are on the right trail as we go.

