import * as React from 'react';
import TextField from '@mui/material/TextField';
import AdapterDateFns from '@mui/lab/AdapterDateFns';
import LocalizationProvider from '@mui/lab/LocalizationProvider';
import DatePicker from '@mui/lab/DatePicker';
import Stack from '@mui/material/Stack';
import Button from '@mui/material/Button';
import Box from "@mui/material/Box";
import EventStatistics from "./EventStatistics";

export default function ViewsDatePicker(props) {
    const [startDate, setStartDate] = React.useState(new Date());
    const [endDate, setEndDate] = React.useState(new Date());
    const { findAllEvents, events } = props;

    function convertMonth(month) {
        let monthToNum = new Map([['Jan', 1], ['Feb', 2], ['Mar', 3], ['Apr', 4],
                                  ['May', 5], ['Jun', 6], ['Jul', 7], ['Agu', 8],
                                  ['Sep', 9], ['Oct', 10], ['Nov', 11], ['Dec', 12]]);
        return monthToNum.get(month);
    }
    function convertDate(date) {
        let dateArr = String(date).split(" ");
        return dateArr[3] + '-' + convertMonth(dateArr[1]) + '-' + dateArr[2];
    }

    return (
        <Box>
            <LocalizationProvider dateAdapter={AdapterDateFns}>
                <Stack spacing={3}>
                    <DatePicker
                        views={['day']}
                        label="start date"
                        value={startDate}
                        onChange={(newValue) => {
                            setStartDate(newValue);
                        }}
                        renderInput={(params) => <TextField {...params} helperText={null} />}
                    />
                    <DatePicker
                        views={['day']}
                        label="end date"
                        value={endDate}
                        onChange={(newValue) => {
                            setEndDate(newValue);
                        }}
                        renderInput={(params) => <TextField {...params} helperText={null} />}
                    />
                
                </Stack>
            </LocalizationProvider>
            <Button size="small" variant="contained" sx={{ mt: 2, ml: 1 }} onClick={() => { findAllEvents(convertDate(startDate), convertDate(endDate)) }}>Submit</Button>
            <Box sx={{mt:2}}>
                <EventStatistics keys={events} />
            </Box>
            
        </Box>
    );
}