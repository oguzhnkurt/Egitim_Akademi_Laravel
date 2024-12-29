import React from 'react';
import ReactDOM from 'react-dom';  
import { createPopper } from '@popperjs/core';
import 'bootstrap';

const button = document.querySelector('#myButton');
const tooltip = document.querySelector('#myTooltip');

createPopper(button, tooltip, {
    placement: 'top',
});

ReactDOM.render(
    <ExamApp />,
    document.getElementById('react-root')  
);
