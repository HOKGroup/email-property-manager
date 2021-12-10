import * as React from "react";
import { styled } from "@mui/material/styles";
import Card from "@mui/material/Card";
import CardContent from "@mui/material/CardContent";
import CardActions from "@mui/material/CardActions";
import IconButton from "@mui/material/IconButton";
import Typography from "@mui/material/Typography";
import Box from "@mui/material/Box";
import Button from "@mui/material/Button";
import AttachEmailIcon from '@mui/icons-material/AttachEmail';
import InsertInvitationIcon from '@mui/icons-material/InsertInvitation';
import DoubleArrowIcon from '@mui/icons-material/DoubleArrow';
import DeleteForeverIcon from '@mui/icons-material/DeleteForever';

export default function ProjectInfoCard(props) {
    console.log(props);
    const { createNewMessage, createAppointment, projectInfo, deleteProject, color } = props;
    return (
        <Card sx={{ maxWidth: 300, boxShadow: 3 }}>
            <Box sx={{ display: "flex", flexDirection: "row" }}>
                <Box sx={{ display: "flex", alignItems: "center", pl: 1 }}>
                    <Button variant="contained" size="small">SET</Button>
                </Box>
                <Box sx={{ display: "flex", flexDirection: "column", pl: 1 }}>
                    <CardContent >
                        <Typography variant="body1" color="text.primary">
                            { projectInfo }
                        </Typography>
                    </CardContent>
                    <CardActions>
                        <Button variant="outlined" size="small">ALT+1</Button>
                        <IconButton aria-label="createNewMessage" sx={{ pl: 1 }} onClick={() => { createNewMessage(projectInfo); }}>
                            <AttachEmailIcon color="primary"/>
                        </IconButton>
                        <IconButton aria-label="creatNewCalendar" onClick={() => { createAppointment(projectInfo + '_') }}>
                            <InsertInvitationIcon color="primary" />
                        </IconButton>
                        <IconButton aria-label="to be determined" onClick={() => {deleteProject(projectInfo + '_' + color); }}>
                            <DeleteForeverIcon color="primary" />
                        </IconButton>
                    </CardActions>
                </Box>
            </Box>
            <div style={{ backgroundColor: color }} height={"5px"}>
                &nbsp;
      </div>
        </Card>
    );
}