The script creates simulated data for the bookings of a hotel and stores them in the excel file. The hotel has lont term guests from 14 -40 days. Half of the groupings are group bookings.
The managers task is to optimize the cleaning plan along the follwoing conditions: 
a) the cleaning should take place as close as possibe to a target date, which is 10 or 20 days from check-in depending whether it is a mid-term or a long term stay. 
b) The maximum number of rooms that can be cleaned in one day is five, since the cleaning personal also has to clean the rooms for the short term guests on a less flexible schedule. 

The script calculates the target day and optimizes the workload as close as possible around the target days. It also puts out the deviation from the target day in the excel file for the managers inspection. 
The optimization is updated dynamically each time a new booking is registerd in the system.  Further automatic processing is now possible, such as automatic e-mail notofication of guests who have an upcoming 
cleaning and information of the cleaning personal which rooms to clean at what dates.
