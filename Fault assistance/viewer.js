import * as THREE from 'https://unpkg.com/three@0.165.0/build/three.module.js';
import { OrbitControls } from 'https://unpkg.com/three@0.165.0/examples/jsm/controls/OrbitControls.js';
import { GLTFLoader } from 'https://unpkg.com/three@0.165.0/examples/jsm/loaders/GLTFLoader.js';

// 🔹 container
const container = document.getElementById('viewer3d');

// 🔹 Scene
const scene = new THREE.Scene();
scene.background = new THREE.Color(0x111111);

// 🔹 Camera
const camera = new THREE.PerspectiveCamera(
  60,
  container.clientWidth / container.clientHeight,
  0.1,
  1000
);
camera.position.set(3, 3, 3);

// 🔹 Renderer
const renderer = new THREE.WebGLRenderer({ antialias: true });
renderer.setSize(container.clientWidth, container.clientHeight);
container.appendChild(renderer.domElement);

// 🔹 Controls (หมุน / Zoom / Pan ✅)
const controls = new OrbitControls(camera, renderer.domElement);
controls.enableDamping = true;

// 🔹 Light
const light1 = new THREE.HemisphereLight(0xffffff, 0x444444);
scene.add(light1);

const light2 = new THREE.DirectionalLight(0xffffff);
light2.position.set(5, 5, 5);
scene.add(light2);

// 🔹 Load GLB model ✅
const loader = new GLTFLoader();
loader.load('./models/2633475_00_Rollenbahn01.glb', (gltf) => {

    scene.add(gltf.scene);

    // Auto fit view
    const box = new THREE.Box3().setFromObject(gltf.scene);
    const center = box.getCenter(new THREE.Vector3());
    const size = box.getSize(new THREE.Vector3()).length();

    camera.position.set(center.x + size, center.y + size, center.z + size);
    controls.target.copy(center);

});

// 🔹 Animate
function animate() {
    requestAnimationFrame(animate);
    controls.update();
    renderer.render(scene, camera);
}
animate();

// 🔹 Resize
window.addEventListener('resize', () => {
    camera.aspect = container.clientWidth / container.clientHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(container.clientWidth, container.clientHeight);
});
