import * as React from 'react';
import Box from '@mui/material/Box';
import Button from '@mui/material/Button';
import Typography from '@mui/material/Typography';
import Modal from '@mui/material/Modal';
import AddCircleRoundedIcon from '@mui/icons-material/AddCircleRounded';
import TextField from "@mui/material/TextField";
import { GithubPicker } from "react-color";

const style = {
  position: 'absolute',
  top: '50%',
  left: '50%',
  transform: 'translate(-50%, -50%)',
  width: 300,
  bgcolor: 'background.paper',
  boxShadow: 3,
  p: 4,
};

export default function BasicModal(props) {
  const addProjectInfo = props.addProjectInfo;
  const [open, setOpen] = React.useState(false);
  const [key, setKey] = React.useState("");
  const [val, setVal] = React.useState("");
  const [color, setColor] = React.useState("#bed3f3");
  const handleOpen = () => setOpen(true);
  const handleClose = () => {
      setOpen(false);
      setKey("");
      setVal("");
    };

  return (
    <div>
      <AddCircleRoundedIcon sx={{ color: 'white', position: 'absolute', right: 5 }} onClick={ handleOpen } />
      <Modal
        open={open}
        onClose={handleClose}
        aria-labelledby="modal-modal-title"
        aria-describedby="modal-modal-description"
      >
        <Box sx={style}>
          <Typography id="modal-modal-title" variant="h6" component="h2"> Add New Project Info </Typography>
          <Typography id="project-name" sx={{ mt: 2 }}> Project Name: </Typography>
          <TextField id="project-name-input" variant="filled" size="small" value={key} onChange={(e) => { setKey(e.target.value) }} />
          <Typography id="project-number" mt={2}> Project Name: </Typography>
          <TextField id="project-number-input" variant="filled" size="small" value={val} onChange={(e) => { setVal(e.target.value) }} />
          <Box sx={{ display: "flex", flexDirection: "row", alignItems: "center", mt: 2 }}>
              <Typography id="selected-color"> Selected Color Label: </Typography>
              <Box sx={{ width: "50px", height: "20px", ml: 1, backgroundColor:color}} />
          </Box>
                  <GithubPicker sx={{ width: "200px", height: "300px", backgroundColor: color}} onChangeComplete={(c) => { setColor(c.hex) }} />
                  <Button
                      onClick={() => {
                      addProjectInfo(key, val, color);
                      handleClose();}}
                      sx={{ mt: 2 }}
                      variant="contained"
                      size="small" >Submit</Button>
        </Box>
      </Modal>
    </div>
  );
}