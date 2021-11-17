import * as React from 'react';
import FormGroup from '@mui/material/FormGroup';
import FormControlLabel from '@mui/material/FormControlLabel';
import Switch from '@mui/material/Switch';

export default function SwitchLabels(props) {
  const switchMode = props.switchMode;
  return (
    <FormGroup>
          <FormControlLabel control={<Switch defaultunchecked="true" />}
              onChange={() => { switchMode(); }}
              label="Read Mode" />
    </FormGroup>
  );
}