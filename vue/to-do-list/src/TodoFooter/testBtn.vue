<script setup lang="ts">
const btnTittle = 'test'
const btnID = 'testBtn'
</script>
<template>
    <div id="learnDiv">
        <button v-bind:id="btnID"> {{ btnTittle }} </button>
    </div>
   
</template>